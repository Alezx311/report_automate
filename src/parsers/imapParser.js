const Imap = require('imap')
const { simpleParser } = require('mailparser')

class ImapParser {
  constructor(config) {
    this.config = {
      user: config.user,
      password: config.password,
      host: config.host || 'outlook.office365.com',
      port: config.port || 993,
      tls: config.tls !== false,
      tlsOptions: { rejectUnauthorized: false },
      connTimeout: 30000,
      authTimeout: 30000,
    }
    this.imap = null
  }

  async connect() {
    return new Promise((resolve, reject) => {
      this.imap = new Imap(this.config)

      this.imap.once('ready', () => {
        console.log('✅ IMAP підключено')
        resolve()
      })

      this.imap.once('error', err => {
        console.error('❌ IMAP помилка:', err)
        reject(err)
      })

      this.imap.connect()
    })
  }

  async disconnect() {
    if (this.imap) {
      this.imap.end()
      this.imap = null
    }
  }

  async listFolders() {
    return new Promise((resolve, reject) => {
      this.imap.getBoxes((err, boxes) => {
        if (err) {
          reject(err)
          return
        }

        const folders = this._flattenBoxes(boxes)
        resolve(folders)
      })
    })
  }

  async fetchEmails(options = {}) {
    const { folder = 'INBOX', startDate, endDate } = options

    return new Promise((resolve, reject) => {
      this.imap.openBox(folder, true, (err, box) => {
        if (err) {
          reject(err)
          return
        }

        console.log(`📁 Папка: ${folder}, повідомлень: ${box.messages.total}`)

        // Формуємо критерії пошуку
        let criteria = ['ALL']

        if (startDate) {
          const start = new Date(startDate)
          criteria.push(['SINCE', start])
        }

        if (endDate) {
          const end = new Date(endDate)
          criteria.push(['BEFORE', end])
        }

        this.imap.search(criteria, (err, uids) => {
          if (err) {
            reject(err)
            return
          }

          if (uids.length === 0) {
            console.log('⚠️ Не знайдено повідомлень')
            resolve([])
            return
          }

          console.log(`📧 Знайдено ${uids.length} повідомлень`)

          const messages = []
          const fetch = this.imap.fetch(uids, {
            bodies: '',
            struct: true,
          })

          let processed = 0

          fetch.on('message', msg => {
            msg.on('body', stream => {
              simpleParser(stream, (err, parsed) => {
                if (err) {
                  console.error('⚠️ Помилка парсингу:', err)
                  return
                }

                messages.push({
                  conversationId: parsed.messageId || parsed.subject || 'unknown',
                  subject: parsed.subject || 'No Subject',
                  senderEmail: parsed.from?.value?.[0]?.address || '',
                  senderName: parsed.from?.value?.[0]?.name || 'Unknown',
                  receivedDateTime: parsed.date || new Date(),
                  body: parsed.text || '',
                  messageId: parsed.messageId,
                  inReplyTo: parsed.inReplyTo,
                  folderName: folder,
                })

                processed++
                if (processed % 10 === 0) {
                  console.log(`  Оброблено ${processed}/${uids.length}`)
                }
              })
            })
          })

          fetch.once('error', err => {
            console.error('❌ Помилка fetch:', err)
            reject(err)
          })

          fetch.once('end', () => {
            console.log(`✅ Завершено. Оброблено ${messages.length} повідомлень`)
            resolve(messages)
          })
        })
      })
    })
  }

  _flattenBoxes(boxes, prefix = '') {
    let result = []

    for (const [name, box] of Object.entries(boxes)) {
      const fullName = prefix ? `${prefix}${box.delimiter}${name}` : name

      result.push({
        name: fullName,
        delimiter: box.delimiter,
        hasChildren: box.children !== null,
        attribs: box.attribs || [],
      })

      if (box.children) {
        result = result.concat(this._flattenBoxes(box.children, fullName))
      }
    }

    return result
  }
}

module.exports = ImapParser
