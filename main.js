const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const fs = require('fs')
require('dotenv').config()

// Парсери та клієнти
const PSTParser = require('./src/parsers/pstParser')
const GraphApiParser = require('./src/parsers/graphApiParser')
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
// Вибір PST файлу або JSON
// ============================================

ipcMain.handle('select-pst-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'PST Files', extensions: ['pst'] },
      { name: 'JSON Files (processed chunks)', extensions: ['json'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  })

  if (!result.canceled && result.filePaths.length > 0) {
    return { success: true, path: result.filePaths[0] }
  }
  return { success: false }
})

// ============================================
// Парсинг з PST або JSON
// ============================================

ipcMain.handle('parse-pst', async (event, options) => {
  try {
    const filePath = options.pstPath
    const fileExt = path.extname(filePath).toLowerCase()

    // Перевіряємо чи це JSON файл з обробленими даними
    if (fileExt === '.json') {
      console.log('Loading processed JSON file...')
      const jsonData = JSON.parse(fs.readFileSync(filePath, 'utf8'))

      // Перевіряємо чи це файл з об'єднаними результатами
      if (jsonData.issues && jsonData.stats) {
        console.log(`Loaded ${jsonData.issues.length} issues from processed JSON`)
        return {
          success: true,
          data: jsonData.issues,
          stats: jsonData.stats,
        }
      }

      // Або це окремий chunk з повідомленнями
      if (jsonData.messages) {
        console.log(`Processing ${jsonData.messages.length} messages from JSON chunk`)
        const reportGenerator = new ReportGenerator({
          supportEmails: options.supportEmails,
          keywords: options.keywords,
        })
        const { issues, stats } = reportGenerator.processMessages(jsonData.messages)
        return {
          success: true,
          data: issues,
          stats,
        }
      }

      throw new Error('Невідомий формат JSON файлу')
    }

    // Інакше це PST файл - стандартна обробка
    console.log('Початок парсингу PST файлу...')

    const pstParser = new PSTParser(options.pstPath)

    const messages = await pstParser.extractMessages({
      startDate: options.startDate,
      endDate: options.endDate,
      batchSize: options.batchSize || 100,
    })

    const reportGenerator = new ReportGenerator({
      supportEmails: options.supportEmails,
      keywords: options.keywords,
      useAggressiveClean: options.useAggressiveClean || false,
    })

    const { issues, stats } = reportGenerator.processMessages(messages)

    return {
      success: true,
      data: issues,
      stats,
    }
  } catch (error) {
    console.error('File parsing error:', error)
    console.error('DEBUG: Stack trace:', error.stack)
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
    console.log('Connecting to Jira...')

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
    console.error('Jira error:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})

ipcMain.handle('fetch-jira-issues', async (event, options) => {
  try {
    console.log('Loading issues from Jira...')

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

    console.log(`Received ${jiraIssues.length} issues`)

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
    console.error('Jira loading error:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})

ipcMain.handle('export-to-jira', async (event, options) => {
  try {
    console.log('Exporting to Jira...')

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
    console.error('Jira export error:', error)
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
    console.error('CSV export error:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})


// ============================================
// Microsoft Graph API - Тестування підключення
// ============================================

ipcMain.handle('test-graph-connection', async (event, credentials) => {
  try {
    console.log('Testing Graph API connection...')
    const graphParser = new GraphApiParser({
      tenant: credentials.tenant || process.env.AZURE_TENANT_ID,
      clientId: credentials.clientId || process.env.AZURE_CLIENT_ID,
      clientSecret: credentials.clientSecret || process.env.AZURE_CLIENT_SECRET,
      user: credentials.user || process.env.GRAPH_API_USER,
      password: credentials.password || process.env.GRAPH_API_PASSWORD,
    })

    const result = await graphParser.testConnection()
    return result
  } catch (error) {
    console.error('Graph API test error:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})

// ============================================
// Microsoft Graph API - Отримання списку папок
// ============================================

ipcMain.handle('get-graph-folders', async (event, credentials) => {
  try {
    console.log('Getting Graph API folders...')
    const graphParser = new GraphApiParser({
      tenant: credentials.tenant || process.env.AZURE_TENANT_ID,
      clientId: credentials.clientId || process.env.AZURE_CLIENT_ID,
      clientSecret: credentials.clientSecret || process.env.AZURE_CLIENT_SECRET,
      user: credentials.user || process.env.GRAPH_API_USER,
      password: credentials.password || process.env.GRAPH_API_PASSWORD,
    })

    const folders = await graphParser.getFoldersWithStats()

    return {
      success: true,
      folders,
    }
  } catch (error) {
    console.error('Graph API folders error:', error)
    return {
      success: false,
      error: error.message,
    }
  }
})

// ============================================
// Microsoft Graph API - Парсинг листів з вибраних папок
// ============================================

ipcMain.handle('parse-graph', async (event, options) => {
  try {
    console.log('Starting Graph API parsing...')
    const graphParser = new GraphApiParser({
      tenant: options.tenant || process.env.AZURE_TENANT_ID,
      clientId: options.clientId || process.env.AZURE_CLIENT_ID,
      clientSecret: options.clientSecret || process.env.AZURE_CLIENT_SECRET,
      user: options.user || process.env.GRAPH_API_USER,
      password: options.password || process.env.GRAPH_API_PASSWORD,
    })

    await graphParser.connect()

    // Витягуємо листи з кожної вибраної папки
    let allMessages = []
    for (const folder of options.folders) {
      console.log(`Fetching from folder: ${folder}`)
      const messages = await graphParser.fetchEmails({
        folder,
        startDate: options.startDate,
        endDate: options.endDate,
      })
      allMessages = allMessages.concat(messages)
    }

    await graphParser.disconnect()

    console.log(`Total messages fetched: ${allMessages.length}`)

    // Генеруємо звіт
    const reportGenerator = new ReportGenerator({
      supportEmails: options.supportEmails,
      keywords: options.keywords,
    })

    const result = reportGenerator.processMessages(allMessages)

    return {
      success: true,
      data: result.issues,
      stats: result.stats,
    }
  } catch (error) {
    console.error('Graph API parse error:', error)
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
        jiraHost: process.env.JIRA_HOST || '',
        jiraProject: process.env.JIRA_PROJECT_KEY || 'SUPPORT',
        jiraEmail: process.env.JIRA_EMAIL || '',
        // Graph API
        azureTenant: process.env.AZURE_TENANT_ID || '',
        azureClientId: process.env.AZURE_CLIENT_ID || '',
        graphUser: process.env.GRAPH_API_USER || '',
      },
    }
  } catch (error) {
    return {
      success: false,
      error: error.message,
    }
  }
})
