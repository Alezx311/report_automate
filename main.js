const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const fs = require('fs')
require('dotenv').config()

// Парсери та клієнти
const PSTParser = require('./src/parsers/pstParser')
const ImapParser = require('./src/parsers/imapParser')
const JiraClient = require('./src/integrations/jiraClient')
const ReportGenerator = require('./src/processors/reportGenerator')

let mainWindow

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1600,
    height: 1000,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
    },
    icon: path.join(__dirname, 'assets/icon.png'),
  })

  mainWindow.loadFile('ui/index.html')

  mainWindow.webContents.openDevTools()
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

// ============================================
// Вибір PST файлу
// ============================================

ipcMain.handle('select-pst-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'PST Files', extensions: ['pst'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  })

  if (!result.canceled && result.filePaths.length > 0) {
    return { success: true, path: result.filePaths[0] }
  }
  return { success: false }
})

// ============================================
// Парсинг з PST
// ============================================

ipcMain.handle('parse-pst', async (event, options) => {
  try {
    console.log('🔄 Початок парсингу PST файлу...')

    const pstParser = new PSTParser(options.pstPath)
    const messages = await pstParser.extractMessages({
      startDate: options.startDate,
      endDate: options.endDate,
      batchSize: options.batchSize || 100,
    })

    console.log(`📧 Отримано ${messages.length} повідомлень`)

    const reportGenerator = new ReportGenerator({
      supportEmails: options.supportEmails,
      keywords: options.keywords,
    })

    const { issues, stats } = reportGenerator.processMessages(messages)

    return {
      success: true,
      data: issues,
      stats,
    }
  } catch (error) {
    console.error('❌ Помилка парсингу PST:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})

// ============================================
// Підключення до IMAP
// ============================================

ipcMain.handle('connect-imap', async (event, config) => {
  try {
    console.log('🔄 Підключення до IMAP...')

    const imapParser = new ImapParser({
      user: config.user || process.env.OUTLOOK_IMAP_USER,
      password: config.password || process.env.OUTLOOK_IMAP_PASSWORD,
      host: config.host || process.env.OUTLOOK_IMAP_HOST || 'outlook.office365.com',
      port: config.port || process.env.OUTLOOK_IMAP_PORT || 993,
    })

    await imapParser.connect()
    const folders = await imapParser.listFolders()
    await imapParser.disconnect()

    return {
      success: true,
      folders: folders.map(f => f.name),
    }
  } catch (error) {
    console.error('❌ Помилка IMAP:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})

ipcMain.handle('parse-imap', async (event, options) => {
  let imapParser = null

  try {
    console.log('🔄 Завантаження листів через IMAP...')

    imapParser = new ImapParser({
      user: options.user || process.env.OUTLOOK_IMAP_USER,
      password: options.password || process.env.OUTLOOK_IMAP_PASSWORD,
      host: options.host || process.env.OUTLOOK_IMAP_HOST,
      port: options.port || process.env.OUTLOOK_IMAP_PORT,
    })

    await imapParser.connect()

    const messages = await imapParser.fetchEmails({
      folder: options.folder || 'INBOX',
      startDate: options.startDate,
      endDate: options.endDate,
    })

    await imapParser.disconnect()

    console.log(`📧 Отримано ${messages.length} повідомлень`)

    const reportGenerator = new ReportGenerator({
      supportEmails: options.supportEmails,
      keywords: options.keywords,
    })

    const { issues, stats } = reportGenerator.processMessages(messages)

    return {
      success: true,
      data: issues,
      stats,
    }
  } catch (error) {
    console.error('❌ Помилка IMAP парсингу:', error)
    if (imapParser) await imapParser.disconnect()
    return {
      success: false,
      error: error.message,
    }
  }
})

// ============================================
// Робота з Jira
// ============================================

ipcMain.handle('connect-jira', async (event, config) => {
  try {
    console.log('🔄 Підключення до Jira...')

    const jiraClient = new JiraClient({
      host: config.host || process.env.JIRA_HOST,
      email: config.email || process.env.JIRA_EMAIL,
      apiToken: config.apiToken || process.env.JIRA_API_TOKEN,
      projectKey: config.projectKey || process.env.JIRA_PROJECT_KEY,
    })

    const result = await jiraClient.testConnection()

    if (result.success) {
      const project = await jiraClient.getProject()
      return {
        success: true,
        user: result.user,
        project: project?.name || 'Unknown',
      }
    }

    return result
  } catch (error) {
    console.error('❌ Помилка Jira:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})

ipcMain.handle('fetch-jira-issues', async (event, options) => {
  try {
    console.log('🔄 Завантаження задач з Jira...')

    const jiraClient = new JiraClient({
      host: options.host || process.env.JIRA_HOST,
      email: options.email || process.env.JIRA_EMAIL,
      apiToken: options.apiToken || process.env.JIRA_API_TOKEN,
      projectKey: options.projectKey || process.env.JIRA_PROJECT_KEY,
    })

    const jiraIssues = await jiraClient.fetchIssues({
      startDate: options.startDate,
      endDate: options.endDate,
      statuses: ['Assigned', 'In Progress', 'Completed'],
    })

    console.log(`📋 Отримано ${jiraIssues.length} задач`)

    const reportGenerator = new ReportGenerator({
      supportEmails: options.supportEmails,
    })

    const { issues, stats } = reportGenerator.processJiraIssues(jiraIssues)

    return {
      success: true,
      data: issues,
      stats,
    }
  } catch (error) {
    console.error('❌ Помилка завантаження Jira:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})

ipcMain.handle('export-to-jira', async (event, options) => {
  try {
    console.log('🔄 Експорт в Jira...')

    const jiraClient = new JiraClient({
      host: options.host || process.env.JIRA_HOST,
      email: options.email || process.env.JIRA_EMAIL,
      apiToken: options.apiToken || process.env.JIRA_API_TOKEN,
      projectKey: options.projectKey || process.env.JIRA_PROJECT_KEY,
    })

    const results = await jiraClient.createBulkIssues(options.issues, (current, total, result) => {
      mainWindow.webContents.send('jira-export-progress', {
        current,
        total,
        result,
      })
    })

    return {
      success: true,
      results,
      created: results.filter(r => r.success).length,
      failed: results.filter(r => !r.success).length,
    }
  } catch (error) {
    console.error('❌ Помилка експорту в Jira:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})

// ============================================
// Експорт в CSV
// ============================================

ipcMain.handle('export-csv', async (event, issues) => {
  try {
    const reportGenerator = new ReportGenerator({})
    const csvPath = await reportGenerator.generateCSV(issues)

    return {
      success: true,
      csvPath,
    }
  } catch (error) {
    console.error('❌ Помилка експорту CSV:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})

// ============================================
// Завантаження налаштувань
// ============================================

ipcMain.handle('load-config', async () => {
  try {
    return {
      success: true,
      config: {
        supportEmails: process.env.SUPPORT_EMAILS || '',
        outlookHost: process.env.OUTLOOK_IMAP_HOST || 'outlook.office365.com',
        outlookPort: process.env.OUTLOOK_IMAP_PORT || '993',
        outlookUser: process.env.OUTLOOK_IMAP_USER || '',
        jiraHost: process.env.JIRA_HOST || '',
        jiraProject: process.env.JIRA_PROJECT_KEY || 'SUPPORT',
        jiraEmail: process.env.JIRA_EMAIL || '',
      },
    }
  } catch (error) {
    return {
      success: false,
      error: error.message,
    }
  }
})
