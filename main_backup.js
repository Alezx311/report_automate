const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const fs = require('fs')
const PST = require('pst-parser')
const createCsvWriter = require('csv-writer').createObjectCsvWriter
const os = require('os')
const iconv = require('iconv-lite')

let mainWindow

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
    },
  })

  mainWindow.loadFile('index.html')
  // mainWindow.webContents.openDevTools(); // Розкоментувати для дебагу
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

// Вибір PST файлу
ipcMain.handle('select-pst-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'PST Files', extensions: ['pst'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  })

  if (!result.canceled && result.filePaths.length > 0) {
    return result.filePaths[0]
  }
  return null
})

// Парсинг PST файлу
ipcMain.handle('parse-pst', async (event, options) => {
  const {
    pstPath,
    supportEmails,
    keywords,
    filterText,
    startDate,
    endDate,
    batchSize = 100,
    ignoreInvalidDates = false,
  } = options

  try {
    // Перевірка розміру файлу
    const statsFs = fs.statSync(pstPath)
    const fileSizeGB = statsFs.size / (1024 * 1024 * 1024)

    if (fileSizeGB > 2) {
      return {
        success: false,
        error: 'Файл завеликий (>2GB). Будь ласка, розділіть PST в Outlook або використовуйте менші експорти.',
      }
    }

    console.debug('Початок парсингу PST файлу...')

    // Читаємо PST файл як ArrayBuffer
    const fileBuffer = fs.readFileSync(pstPath)
    const arrayBuffer = fileBuffer.buffer.slice(fileBuffer.byteOffset, fileBuffer.byteOffset + fileBuffer.byteLength)

    // Створюємо PST об'єкт
    const pst = new PST.PSTFile(arrayBuffer)
    const messageStore = pst.getMessageStore()
    const rootFolder = messageStore.getRootFolder()

    console.debug(`Коренева папка: ${rootFolder.displayName}`)

    const messages = []

    const startD = startDate ? new Date(startDate) : null
    const endD = endDate ? new Date(endDate + 'T23:59:59') : null

    // Рекурсивна функція для обробки папок
    function processFolderRecursive(folder, depth = 0) {
      const indent = '  '.repeat(depth)
      console.debug(`${indent} ${fixEncoding(folder.displayName)} (глибина: ${depth})`)

      try {
        // Отримуємо повідомлення з папки
        const messageCount = folder.contentCount || 0
        if (messageCount > 0) {
          console.debug(`${indent} Повідомлень: ${messageCount}`)

          // Отримуємо всі повідомлення (по 100 за раз для продуктивності)
          let offset = 0
          // const batchSize = 100

          while (offset < messageCount) {
            try {
              const messageEntries = folder.getContents(offset, batchSize)

              for (const entry of messageEntries) {
                try {
                  const message = folder.getMessage(entry.nid)

                  if (message) {
                    const rawBody = message.body || message.bodyHTML || ''
                    const rawSubject = message.subject || 'No Subject'

                    // Виправляємо кодування
                    const cleanBody = stripHtml(fixEncoding(rawBody))
                    const cleanSubject = fixEncoding(rawSubject)

                    let date = parseDateFromBody(rawBody) // Витяг з body
                    if (ignoreInvalidDates && date.toISOString().startsWith(new Date().toISOString().slice(0, 10)))
                      continue // Ігнор fallback

                    if (startD && date < startD) continue
                    if (endD && date > endD) continue

                    messages.push({
                      conversationId: fixEncoding(message.conversationTopic || cleanSubject || 'unknown'),
                      subject: cleanSubject,
                      senderEmail: message.senderEmailAddress || message.sentRepresentingEmailAddress || '',
                      senderName: fixEncoding(
                        message.senderName || message.sentRepresentingName || message.displayFrom || 'Unknown',
                      ),
                      receivedDateTime: date,
                      body: cleanBody,
                      bodyPreview: cleanBody.substring(0, 500),
                    })
                  }
                } catch (msgError) {
                  console.error(`${indent} Помилка обробки повідомлення:`, msgError.message)
                }
              }

              offset += batchSize
            } catch (batchError) {
              console.error(`${indent} Помилка отримання batch:`, batchError.message)
              break
            }
          }
        }

        // Обробка підпапок
        if (folder.hasSubfolders) {
          const subFolderEntries = folder.getSubFolderEntries()
          console.debug(`${indent} Підпапок: ${subFolderEntries.length}`)

          for (const entry of subFolderEntries) {
            try {
              const subFolder = folder.getSubFolder(entry.nid)
              if (subFolder) {
                processFolderRecursive(subFolder, depth + 1)
              }
            } catch (subFolderError) {
              console.error(`${indent} Помилка обробки підпапки:`, subFolderError.message)
            }
          }
        }
      } catch (folderError) {
        console.error(`${indent} Помилка обробки папки:`, folderError.message)
      }
    }

    // Почати обробку з кореневої папки
    processFolderRecursive(rootFolder)

    console.debug(`Знайдено повідомлень: ${messages.length}`)

    if (messages.length === 0) {
      return {
        success: false,
        error: 'Немає даних у PST файлі',
      }
    }

    // Групування в threads
    const threads = groupByThread(messages)
    console.debug(`Згруповано в threads: ${Object.keys(threads).length}`)

    // Обробка threads та генерація звернень
    const { issues, stats } = processThreads(threads, supportEmails, keywords, filterText)

    console.debug(`Результат processThreads: issues=${issues.length}, stats=`, stats)

    const csvPath = await generateCSV(issues)

    return {
      success: true,
      data: issues,
      stats,
      csvPath,
    }
  } catch (error) {
    console.error('Критична помилка парсингу:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})

function fixEncoding(str) {
  return str
  if (!str) return ''
  const buffer = Buffer.from(str, 'latin1')
  return iconv.decode(buffer, 'win1251')
}

function stripHtml(html) {
  return html
    .replace(/<[^>]*>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

function groupByThread(messages) {
  const threads = {}
  messages.forEach(msg => {
    const convId = msg.conversationId
    if (!threads[convId]) threads[convId] = []
    threads[convId].push(msg)
  })
  return threads
}

function processThreads(threads, supportEmailsStr, keywordsStr, filterText) {
  const supportEmails = supportEmailsStr
    .split(',')
    .map(e => e.trim().toLowerCase())
    .filter(e => e)
  const keywords = keywordsStr
    .split(',')
    .map(k => k.trim().toLowerCase())
    .filter(k => k)

  const emailToName = {
    'oleksiy_sokolov@service-team.biz': 'Олексій Соколов',
    'dmytro_sandul@service-team.biz': 'Дмитро Сандул',
    'nikita_chychykalo@service-team.biz': 'Нікіта Чичикало',
    'ihor_draha@service-team.biz': 'Ігор Драга',
    // Додай інші email, якщо потрібно
  }

  function getResponsible(email) {
    const lowerEmail = email.toLowerCase()
    for (const [key, name] of Object.entries(emailToName)) {
      if (lowerEmail.includes(key)) return name
    }
    return 'Невідомо'
  }

  const issues = []

  Object.values(threads).forEach(thread => {
    // Фільтр: Якщо filterText, перевіряємо match в будь-якому msg
    if (
      filterText &&
      !thread.some(
        msg =>
          msg.subject.toLowerCase().includes(filterText.toLowerCase()) ||
          msg.body.toLowerCase().includes(filterText.toLowerCase()),
      )
    )
      return

    // Сортуємо thread за часом
    thread.sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime))

    let i = 0
    while (i < thread.length) {
      const currentMsg = thread[i]
      const isSupport = supportEmails.some(email => currentMsg.senderEmail.toLowerCase().includes(email))

      let issue = {
        dateRegistered: formatDate(currentMsg.receivedDateTime),
        timeRegistered: formatTime(currentMsg.receivedDateTime),
        system: extractSystem(currentMsg.body),
        ticketId: '',
        communication: 'Пошта',
        description: extractDescription(currentMsg.body, keywords),
        status: 'У процесі',
        responsible: '',
        solution: '',
        dateResolved: '',
        timeResolved: '',
        importance: calculateImportance(currentMsg.body),
        subject: currentMsg.subject,
      }

      // Знаходимо наступне повідомлення з протилежним типом
      let j = i + 1
      while (j < thread.length) {
        const nextMsg = thread[j]
        const nextIsSupport = supportEmails.some(email => nextMsg.senderEmail.toLowerCase().includes(email))

        if (nextIsSupport !== isSupport) {
          // Це відповідь
          if (isSupport) {
            // Перший - support, наступний - клієнт (відповідь клієнта)
            issue.description = 'Питання від техпідтримки: ' + issue.description
            issue.solution = extractSolution(nextMsg.body)
            issue.status = 'Вирішено' // Або частково, залежно від логіки
            issue.responsible = getResponsible(currentMsg.senderEmail)
            issue.dateResolved = formatDate(nextMsg.receivedDateTime)
            issue.timeResolved = formatTime(nextMsg.receivedDateTime)
          } else {
            // Перший - клієнт, наступний - support (відповідь support)
            issue.responsible = getResponsible(nextMsg.senderEmail)
            issue.solution = extractSolution(nextMsg.body)
            issue.dateResolved = formatDate(nextMsg.receivedDateTime)
            issue.timeResolved = formatTime(nextMsg.receivedDateTime)
            const hasThankYou =
              nextMsg.body.toLowerCase().includes('дякую') || nextMsg.body.toLowerCase().includes('працює')
            issue.status = hasThankYou ? 'Вирішено' : 'Вирішено частково'
          }
          i = j // Перейти до наступного після пари
          break
        }
        j++
      }

      if (j >= thread.length) {
        // Немає пари
        if (isSupport) {
          issue.description = 'Питання від техпідтримки без відповіді: ' + issue.description
          issue.status = 'Не закрито'
          issue.importance = 'Високий'
        } else {
          issue.status = 'Не закрито'
          issue.importance = 'Високий'
        }
      }

      issues.push(issue)
      i++
    }
  })

  const stats = {
    totalThreads: Object.keys(threads).length,
    total: issues.length,
    resolved: issues.filter(i => i.status === 'Вирішено').length,
    partial: issues.filter(i => i.status === 'Вирішено частково').length,
    inProgress: issues.filter(i => i.status === 'У процесі' || i.status === 'Не закрито').length,
  }

  return { issues, stats }
}

function formatDate(dateTime) {
  const date = new Date(dateTime)
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

function formatTime(dateTime) {
  const date = new Date(dateTime)
  const hours = String(date.getHours()).padStart(2, '0')
  const minutes = String(date.getMinutes()).padStart(2, '0')
  const seconds = String(date.getSeconds()).padStart(2, '0')
  return `${hours}:${minutes}:${seconds}`
}

function extractSystem(body) {
  const match = body.match(/([A-Z]+_[A-Z]+)/g) // Витяг uppercase системи, як "IPS"
  return match ? match[0] : '' // Порожнє для чорновика
}

function extractDescription(body, keywords) {
  let description = stripHtml(body).substring(0, 1000)
  const foundKeywords = keywords.filter(keyword => body.toLowerCase().includes(keyword.toLowerCase()))
  if (foundKeywords.length > 0) {
    description += ' [Ключові слова: ' + foundKeywords.join(', ') + ']'
  }
  return description
}

function extractSolution(body) {
  return stripHtml(body).substring(0, 500)
}

function calculateImportance(body) {
  const highKeywords = ['терміново', 'критично', 'urgent', 'critical', 'високий', 'прод']
  const lowKeywords = ['тест', 'аналіз', 'консультація']
  if (highKeywords.some(kw => body.toLowerCase().includes(kw))) return 'Високий'
  if (lowKeywords.some(kw => body.toLowerCase().includes(kw))) return 'Низький'
  return 'Середній'
}

// функція для парсингу дати з body
function parseDateFromBody(body) {
  const datePattern = /Sent:\s*(\w+day),\s*(\w+)\s*(\d+),\s*(\d+)\s*(\d+:\d+\s*[AP]M)/i
  const match = body.match(datePattern)
  if (match) {
    const [, dayOfWeek, month, day, year, time] = match
    const dateStr = `${month} ${day}, ${year} ${time}`
    const parsed = Date.parse(dateStr)
    if (!isNaN(parsed)) {
      return new Date(parsed)
    }
  }
  return null
}

function generateCSV(issues) {
  const outputDir = path.join(__dirname, 'output')

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true })
  }

  const fileName = `report_draft_${formatDate(new Date())}.csv`
  const filePath = path.join(outputDir, fileName)

  const csvWriter = createCsvWriter({
    path: filePath,
    header: [
      { id: 'dateRegistered', title: 'Дата реєстрації проблеми' },
      { id: 'timeRegistered', title: 'Час реєстрації проблеми (приблизно)' },
      { id: 'system', title: 'Система' },
      { id: 'ticketId', title: 'ID заявки у SD' },
      { id: 'communication', title: 'Комунікація' },
      { id: 'description', title: 'Опис проблеми' },
      { id: 'status', title: 'Статус' },
      { id: 'responsible', title: 'Відповідальний' },
      { id: 'solution', title: 'Рішення' },
      { id: 'dateResolved', title: 'Дата вирішення проблеми' },
      { id: 'timeResolved', title: 'Час вирішення проблеми (приблизно)' },
      { id: 'importance', title: 'Критерій важливості' },
    ],
    encoding: 'utf8',
  })

  csvWriter.writeRecords(issues).then(() => console.debug(`CSV створено: ${filePath}`))
  return filePath
}
