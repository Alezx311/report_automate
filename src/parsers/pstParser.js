const fs = require('fs')
const PST = require('pst-parser')

class PSTParser {
  constructor(pstPath) {
    this.pstPath = pstPath
    this.pst = null
  }

  async extractMessages(options = {}) {
    const { startDate, endDate, batchSize = 100 } = options

    // Перевірка розміру файлу
    const stats = fs.statSync(this.pstPath)
    const fileSizeGB = stats.size / (1024 * 1024 * 1024)

    if (fileSizeGB > 2) {
      throw new Error('Файл завеликий (>2GB). Розділіть PST в Outlook.')
    }

    console.log(`📂 Читання PST файлу (${fileSizeGB.toFixed(2)} GB)...`)

    // Читаємо файл
    const fileBuffer = fs.readFileSync(this.pstPath)
    const arrayBuffer = fileBuffer.buffer.slice(fileBuffer.byteOffset, fileBuffer.byteOffset + fileBuffer.byteLength)

    this.pst = new PST.PSTFile(arrayBuffer)
    const messageStore = this.pst.getMessageStore()
    const rootFolder = messageStore.getRootFolder()

    const messages = []
    const startD = startDate ? new Date(startDate) : null
    const endD = endDate ? new Date(endDate + 'T23:59:59') : null

    // Рекурсивна обробка папок
    const processFolder = (folder, depth = 0) => {
      const indent = '  '.repeat(depth)
      console.log(`${indent}📁 ${folder.displayName}`)

      try {
        const messageCount = folder.contentCount || 0

        if (messageCount > 0) {
          console.log(`${indent}  📧 ${messageCount} повідомлень`)

          let offset = 0
          while (offset < messageCount) {
            try {
              const entries = folder.getContents(offset, batchSize)

              for (const entry of entries) {
                try {
                  const message = folder.getMessage(entry.nid)

                  if (message && message.messageClass === 'IPM.Note') {
                    const date = message.messageDeliveryTime || message.clientSubmitTime || new Date()

                    // Фільтр по датах
                    if (startD && date < startD) continue
                    if (endD && date > endD) continue

                    messages.push({
                      conversationId: message.conversationTopic || message.subject || 'unknown',
                      subject: message.subject || 'No Subject',
                      senderEmail: message.senderEmailAddress || message.sentRepresentingEmailAddress || '',
                      senderName: message.senderName || message.sentRepresentingName || 'Unknown',
                      receivedDateTime: date,
                      body: this.stripHtml(message.body || message.bodyHTML || ''),
                      folderName: folder.displayName,
                    })
                  }
                } catch (msgError) {
                  console.error(`${indent}  ⚠️ Помилка повідомлення:`, msgError.message)
                }
              }

              offset += batchSize
            } catch (batchError) {
              console.error(`${indent}  ⚠️ Помилка batch:`, batchError.message)
              break
            }
          }
        }

        // Обробка підпапок
        if (folder.hasSubfolders) {
          const subFolders = folder.getSubFolderEntries()

          for (const entry of subFolders) {
            try {
              const subFolder = folder.getSubFolder(entry.nid)
              if (subFolder) {
                processFolder(subFolder, depth + 1)
              }
            } catch (subError) {
              console.error(`${indent}  ⚠️ Помилка підпапки:`, subError.message)
            }
          }
        }
      } catch (folderError) {
        console.error(`${indent}⚠️ Помилка папки:`, folderError.message)
      }
    }

    processFolder(rootFolder)

    console.log(`✅ Завершено. Знайдено ${messages.length} повідомлень`)

    return messages
  }

  stripHtml(html) {
    if (!html) return ''

    return html
      .replace(/<[^>]*>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/\s+/g, ' ')
      .trim()
  }
}

module.exports = PSTParser
