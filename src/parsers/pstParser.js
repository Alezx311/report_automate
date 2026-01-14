const fs = require('fs')
const PST = require('pst-parser')
const TextCleaner = require('../utils/textCleaner')

class PSTParser {
  constructor(pstPath) {
    this.pstPath = pstPath
    this.pst = null
  }

  async extractMessages(options = {}) {
    const { startDate, endDate } = options

    // File size check
    const stats = fs.statSync(this.pstPath)
    const fileSizeGB = stats.size / (1024 * 1024 * 1024)

    if (fileSizeGB > 2) {
      throw new Error('File too large (>2GB). Split PST in Outlook.')
    }

    console.log(`Reading PST file (${fileSizeGB.toFixed(2)} GB)...`)

    // Read file
    const fileBuffer = fs.readFileSync(this.pstPath)
    const arrayBuffer = fileBuffer.buffer.slice(fileBuffer.byteOffset, fileBuffer.byteOffset + fileBuffer.byteLength)

    this.pst = new PST.PSTFile(arrayBuffer)
    const messageStore = this.pst.getMessageStore()
    const rootFolder = messageStore.getRootFolder()

    const messages = []
    const startD = startDate ? new Date(startDate) : null
    const endD = endDate ? new Date(endDate + 'T23:59:59') : null

    // Recursive folder processing
    const processFolder = (folder, depth = 0) => {
      const indent = '  '.repeat(depth)
      console.log(`${indent}Folder: ${folder.displayName}`)

      try {
        const messageCount = folder.contentCount || 0

        if (messageCount > 0) {
          console.log(`${indent}  Messages: ${messageCount}`)

          let processedInFolder = 0
          let skippedByDate = 0
          let skippedByClass = 0

          try {
            // Read ALL entries at once (getContents with offset doesn't work properly in pst-parser)
            console.log(`${indent}  Reading all ${messageCount} entries...`)
            const allEntries = folder.getContents(0, messageCount)
            console.log(`${indent}  Got ${allEntries.length} entries`)

            // Process entries
            for (const entry of allEntries) {
              try {
                const message = folder.getMessage(entry.nid)

                if (message) {
                  // Accept message if it's IPM.Note OR if messageClass is undefined but has subject/body
                  const isEmailMessage =
                    message.messageClass === 'IPM.Note' ||
                    (message.messageClass === undefined && (message.subject || message.body || message.bodyHTML))

                  if (isEmailMessage) {
                    // Try to extract date from body text
                    let date = this.extractDateFromBody(message.body, message.bodyHTML)

                    if (!date) {
                      // Fallback to current date if extraction failed
                      date = new Date()
                    }

                    // Date filter
                    if (startD && date < startD) {
                      skippedByDate++
                      continue
                    }
                    if (endD && date > endD) {
                      skippedByDate++
                      continue
                    }

                    // Використовуємо базове очищення для отримання тексту з HTML
                    const bodyText = TextCleaner.basicClean(message.body || message.bodyHTML || '')

                    // Extract email from body if not available in message properties
                    let senderEmail = message.senderEmailAddress || message.sentRepresentingEmailAddress || ''
                    if (!senderEmail || senderEmail.trim() === '') {
                      senderEmail = this.extractEmailFromBody(bodyText)
                    }

                    const msg = {
                      conversationId: message.conversationTopic || message.subject || 'unknown',
                      subject: message.subject || 'No Subject',
                      senderEmail: senderEmail,
                      senderName: message.senderName || message.sentRepresentingName || 'Unknown',
                      receivedDateTime: date,
                      body: bodyText,
                      folderName: folder.displayName,
                    }
                    messages.push(msg)
                    processedInFolder++
                  } else {
                    skippedByClass++
                  }
                }
              } catch (msgError) {
                console.error(`${indent}  Message error:`, msgError.message)
              }
            }

            console.log(`${indent}  Processed: ${processedInFolder}, Skipped by date: ${skippedByDate}, Skipped by type: ${skippedByClass}`)
          } catch (entriesError) {
            console.error(`${indent}  Error reading entries:`, entriesError.message)
          }
        }

        // Process subfolders
        if (folder.hasSubfolders) {
          const subFolders = folder.getSubFolderEntries()

          for (const entry of subFolders) {
            try {
              const subFolder = folder.getSubFolder(entry.nid)
              if (subFolder) {
                processFolder(subFolder, depth + 1)
              }
            } catch (subError) {
              console.error(`${indent}  Subfolder error:`, subError.message)
            }
          }
        }
      } catch (folderError) {
        console.error(`${indent}Folder error:`, folderError.message)
      }
    }

    processFolder(rootFolder)

    console.log(`Завершено. Знайдено ${messages.length} повідомлень`)

    return messages
  }

  extractDateFromBody(body, bodyHTML) {
    // Try to extract date from body text
    const text = body || bodyHTML || ''

    // Try multiple patterns to find the date
    const patterns = [
      // Pattern 1: "Sent: Friday, November 29, 2024 2:04 PM"
      /Sent:\s+\w+,\s+(\w+\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}\s+(?:AM|PM))/i,

      // Pattern 2: "From:... Sent: November 29, 2024 2:04 PM" (without day of week)
      /Sent:\s+(\w+\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}\s+(?:AM|PM))/i,

      // Pattern 3: Look for any date in format "Month DD, YYYY HH:MM AM/PM"
      /(\w+\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}\s+(?:AM|PM))/i,

      // Pattern 4: Date in format "DD/MM/YYYY" or "MM/DD/YYYY"
      /(\d{1,2}\/\d{1,2}\/\d{4})/,
    ]

    for (const pattern of patterns) {
      const match = text.match(pattern)

      if (match) {
        try {
          const dateStr = match[1]
          const parsedDate = new Date(dateStr)

          if (!isNaN(parsedDate.getTime())) {
            return parsedDate
          }
        } catch (e) {
          // Try next pattern
          continue
        }
      }
    }

    return null
  }

  extractEmailFromBody(body) {
    if (!body) return ''

    // Стратегія: Шукаємо email в ПІДПИСІ, а не в forwarded headers
    // Підпис зазвичай знаходиться ПЕРЕД "From:" header (якщо лист forwarded)

    // Розділяємо на частини по "From:" (щоб відокремити forwarded headers)
    const parts = body.split(/From:/i)

    // Якщо є forwarded content (більше 1 частини), беремо ПЕРШУ частину (до headers)
    // Це буде містити підпис РЕАЛЬНОГО відправника
    const actualContent = parts[0]

    // Pattern 1: Email в підписі з вертикальною рискою (наприклад: "Name | email@domain.com")
    const signaturePattern1 = /\|\s*([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/
    const sig1Match = actualContent.match(signaturePattern1)
    if (sig1Match) return sig1Match[1]

    // Pattern 2: Email в <mailto:...> тегу (часто в підписах)
    const mailtoPattern = /mailto:([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/i
    const mailtoMatch = actualContent.match(mailtoPattern)
    if (mailtoMatch) return mailtoMatch[1]

    // Pattern 3: Email біля імені в кінці (підпис)
    // Шукаємо ОСТАННІЙ email в actualContent (найімовірніше це підпис)
    const allEmailsPattern = /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/g
    const allEmails = actualContent.match(allEmailsPattern)
    if (allEmails && allEmails.length > 0) {
      // Повертаємо ОСТАННІЙ знайдений email (найімовірніше з підпису)
      return allEmails[allEmails.length - 1]
    }

    return ''
  }
}

module.exports = PSTParser
