const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const fs = require('fs')
const PST = require('pst-parser')
const createCsvWriter = require('csv-writer').createObjectCsvWriter
const os = require('os')
const iconv = require('iconv-lite')
const utils = require('./utils')

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
      console.debug(`${indent} ${utils.fixEncoding(folder.displayName)} (глибина: ${depth})`)

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
                    const messageBody = utils.stripHtml(utils.fixEncoding(rawBody))
                    const cleanSubject = utils.fixEncoding(rawSubject)

                    let date = utils.parseDateFromBody(rawBody) // Витяг з body
                    if (ignoreInvalidDates && date.toISOString().startsWith(new Date().toISOString().slice(0, 10)))
                      continue // Ігнор fallback

                    if (startD && date < startD) continue
                    if (endD && date > endD) continue

                    messages.push({
                      conversationId: utils.fixEncoding(message.conversationTopic || cleanSubject || 'unknown'),
                      subject: cleanSubject,
                      senderEmail: message.senderEmailAddress || message.sentRepresentingEmailAddress || '',
                      senderName: utils.fixEncoding(
                        message.senderName || message.sentRepresentingName || message.displayFrom || 'Unknown',
                      ),
                      receivedDateTime: date,
                      body: utils.cleanBody(messageBody),
                      bodyPreview: messageBody.substring(0, 500),
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
    const threads = utils.groupByThread(messages)
    console.debug(`Згруповано в threads: ${Object.keys(threads).length}`)

    // Обробка threads та генерація звернень
    const { issues, stats } = utils.processThreads(threads, supportEmails, keywords, filterText)

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

function generateCSV(issues) {
  const outputDir = path.join(__dirname, 'output')

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true })
  }

  const fileName = `report_draft_${utils.formatDate(new Date())}.csv`
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
