// Глобальні змінні
let parsedData = null
let selectedSource = 'pst' // 'pst', 'imap', 'jira'

// Елементи UI
const sourceButtons = {
  pst: document.getElementById('source-pst'),
  imap: document.getElementById('source-imap'),
  jira: document.getElementById('source-jira'),
}

const sourcePanels = {
  pst: document.getElementById('pst-panel'),
  imap: document.getElementById('imap-panel'),
  jira: document.getElementById('jira-panel'),
}

const parseBtn = document.getElementById('parse-btn')
const exportCsvBtn = document.getElementById('export-csv-btn')
const exportJiraBtn = document.getElementById('export-jira-btn')

const resultSection = document.getElementById('result-section')
const resultInfo = document.getElementById('result-info')
const previewSection = document.getElementById('preview-section')
const previewBody = document.getElementById('preview-body')
const loading = document.getElementById('loading')

// ============================================
// Ініціалізація
// ============================================

async function init() {
  console.log('Ініціалізація додатка...')

  // Перевіряємо electronAPI
  if (!window.electronAPI) {
    console.error('❌ electronAPI не знайдено!')
    alert('Помилка: electronAPI не доступний. Перезапустіть додаток.')
    return
  }

  console.log('✅ electronAPI доступний:', Object.keys(window.electronAPI))

  // Завантажуємо конфігурацію
  try {
    const result = await window.electronAPI.loadConfig()

    if (result.success) {
      console.log('✅ Конфігурація завантажена:', result.config)

      // Заповнюємо поля з .env
      document.getElementById('support-emails').value = result.config.supportEmails
      document.getElementById('imap-user').value = result.config.outlookUser
      document.getElementById('imap-host').value = result.config.outlookHost
      document.getElementById('imap-port').value = result.config.outlookPort
      document.getElementById('jira-host').value = result.config.jiraHost
      document.getElementById('jira-email').value = result.config.jiraEmail
      document.getElementById('jira-project').value = result.config.jiraProject
    } else {
      console.warn('⚠️ Помилка завантаження конфігурації:', result.error)
    }
  } catch (error) {
    console.error('❌ Помилка ініціалізації:', error)
  }

  setupEventListeners()
}

// ============================================
// Event Listeners
// ============================================

function setupEventListeners() {
  // Вибір джерела
  Object.keys(sourceButtons).forEach(source => {
    sourceButtons[source].addEventListener('click', () => selectSource(source))
  })

  // PST
  document.getElementById('select-pst-btn').addEventListener('click', selectPSTFile)

  // IMAP
  document.getElementById('connect-imap-btn').addEventListener('click', connectIMAP)

  // Jira
  document.getElementById('connect-jira-btn').addEventListener('click', connectJira)

  // Парсинг
  parseBtn.addEventListener('click', startParsing)

  // Експорт
  exportCsvBtn.addEventListener('click', exportToCSV)
  exportJiraBtn.addEventListener('click', exportToJira)

  // Прогрес Jira
  window.electronAPI.onJiraProgress(data => {
    console.log(`Jira прогрес: ${data.current}/${data.total}`)
    updateJiraProgress(data)
  })
}

// ============================================
// Вибір джерела даних
// ============================================

function selectSource(source) {
  selectedSource = source

  // Оновлюємо кнопки
  Object.keys(sourceButtons).forEach(key => {
    if (key === source) {
      sourceButtons[key].classList.add('active')
    } else {
      sourceButtons[key].classList.remove('active')
    }
  })

  // Показуємо відповідну панель
  Object.keys(sourcePanels).forEach(key => {
    sourcePanels[key].style.display = key === source ? 'block' : 'none'
  })

  console.log(`Обрано джерело: ${source}`)
}

// ============================================
// PST
// ============================================

async function selectPSTFile() {
  try {
    const result = await window.electronAPI.selectPSTFile()

    if (result.success) {
      document.getElementById('pst-path').value = result.path
      parseBtn.disabled = false
      console.log('✅ PST файл обрано:', result.path)
    }
  } catch (error) {
    console.error('❌ Помилка вибору PST:', error)
    alert('Помилка вибору файлу: ' + error.message)
  }
}

// ============================================
// IMAP
// ============================================

async function connectIMAP() {
  const btn = document.getElementById('connect-imap-btn')
  const originalText = btn.textContent

  try {
    btn.disabled = true
    btn.textContent = '⏳ Підключення...'

    const config = {
      user: document.getElementById('imap-user').value,
      password: document.getElementById('imap-password').value,
      host: document.getElementById('imap-host').value,
      port: document.getElementById('imap-port').value,
    }

    if (!config.user || !config.password) {
      alert('Введіть email та пароль для IMAP')
      return
    }

    const result = await window.electronAPI.connectIMAP(config)

    if (result.success) {
      console.log('✅ IMAP підключено. Папки:', result.folders)

      // Заповнюємо список папок
      const folderSelect = document.getElementById('imap-folder')
      folderSelect.innerHTML = ''

      result.folders.forEach(folder => {
        const option = document.createElement('option')
        option.value = folder
        option.textContent = folder
        folderSelect.appendChild(option)
      })

      parseBtn.disabled = false
      alert('✅ Успішно підключено до Outlook!\n\nЗнайдено папок: ' + result.folders.length)
    } else {
      alert('❌ Помилка підключення:\n\n' + result.error)
    }
  } catch (error) {
    console.error('❌ Помилка IMAP:', error)
    alert('Помилка: ' + error.message)
  } finally {
    btn.disabled = false
    btn.textContent = originalText
  }
}

// ============================================
// Jira
// ============================================

async function connectJira() {
  const btn = document.getElementById('connect-jira-btn')
  const originalText = btn.textContent

  try {
    btn.disabled = true
    btn.textContent = '⏳ Підключення...'

    const config = {
      host: document.getElementById('jira-host').value,
      email: document.getElementById('jira-email').value,
      apiToken: document.getElementById('jira-token').value,
      projectKey: document.getElementById('jira-project').value,
    }

    if (!config.host || !config.email || !config.apiToken) {
      alert('Заповніть всі поля для Jira')
      return
    }

    const result = await window.electronAPI.connectJira(config)

    if (result.success) {
      console.log('✅ Jira підключено:', result)
      parseBtn.disabled = false
      alert(`✅ Успішно підключено до Jira!\n\nКористувач: ${result.user.displayName}\nПроект: ${result.project}`)
    } else {
      alert('❌ Помилка підключення:\n\n' + result.error)
    }
  } catch (error) {
    console.error('❌ Помилка Jira:', error)
    alert('Помилка: ' + error.message)
  } finally {
    btn.disabled = false
    btn.textContent = originalText
  }
}

// ============================================
// Парсинг
// ============================================

async function startParsing() {
  try {
    loading.style.display = 'block'
    resultSection.style.display = 'none'
    previewSection.style.display = 'none'
    parseBtn.disabled = true

    let result

    if (selectedSource === 'pst') {
      result = await parsePST()
    } else if (selectedSource === 'imap') {
      result = await parseIMAP()
    } else if (selectedSource === 'jira') {
      result = await parseJira()
    }

    loading.style.display = 'none'
    parseBtn.disabled = false

    if (result.success) {
      parsedData = result.data
      displayResults(result)
    } else {
      showError(result.error)
    }
  } catch (error) {
    loading.style.display = 'none'
    parseBtn.disabled = false
    console.error('❌ Помилка парсингу:', error)
    showError(error.message)
  }
}

async function parsePST() {
  const options = {
    pstPath: document.getElementById('pst-path').value,
    supportEmails: document.getElementById('support-emails').value,
    keywords: document.getElementById('keywords').value,
    startDate: document.getElementById('start-date').value,
    endDate: document.getElementById('end-date').value,
    batchSize: 100,
  }

  if (!options.pstPath) {
    throw new Error('Оберіть PST файл')
  }

  return await window.electronAPI.parsePST(options)
}

async function parseIMAP() {
  const options = {
    user: document.getElementById('imap-user').value,
    password: document.getElementById('imap-password').value,
    host: document.getElementById('imap-host').value,
    port: document.getElementById('imap-port').value,
    folder: document.getElementById('imap-folder').value,
    supportEmails: document.getElementById('support-emails').value,
    keywords: document.getElementById('keywords').value,
    startDate: document.getElementById('start-date').value,
    endDate: document.getElementById('end-date').value,
  }

  if (!options.user || !options.password) {
    throw new Error('Введіть credentials для IMAP')
  }

  return await window.electronAPI.parseIMAP(options)
}

async function parseJira() {
  const options = {
    host: document.getElementById('jira-host').value,
    email: document.getElementById('jira-email').value,
    apiToken: document.getElementById('jira-token').value,
    projectKey: document.getElementById('jira-project').value,
    supportEmails: document.getElementById('support-emails').value,
    startDate: document.getElementById('start-date').value,
    endDate: document.getElementById('end-date').value,
  }

  if (!options.host || !options.email || !options.apiToken) {
    throw new Error('Заповніть credentials для Jira')
  }

  return await window.electronAPI.fetchJiraIssues(options)
}

// ============================================
// Відображення результатів
// ============================================

function displayResults(result) {
  const stats = result.stats

  resultSection.style.display = 'block'
  resultInfo.innerHTML = `
    <div class="success-message">
      <strong>✅ Парсинг завершено!</strong><br><br>
      <div class="stats-grid">
        <div class="stat-item">
          <div class="stat-label">Threads:</div>
          <div class="stat-value">${stats.totalThreads}</div>
        </div>
        <div class="stat-item">
          <div class="stat-label">Звернень:</div>
          <div class="stat-value">${stats.total}</div>
        </div>
        <div class="stat-item stat-resolved">
          <div class="stat-label">Вирішено:</div>
          <div class="stat-value">${stats.resolved}</div>
        </div>
        <div class="stat-item stat-progress">
          <div class="stat-label">У процесі:</div>
          <div class="stat-value">${stats.inProgress}</div>
        </div>
        <div class="stat-item">
          <div class="stat-label">Середньо листів:</div>
          <div class="stat-value">${stats.avgMessagesPerIssue}</div>
        </div>
      </div>
    </div>
  `

  previewSection.style.display = 'block'
  displayTable(result.data)
}

function displayTable(issues) {
  previewBody.innerHTML = ''

  issues.forEach((issue, index) => {
    const row = document.createElement('tr')
    row.innerHTML = `
      <td>${issue.dateRegistered}</td>
      <td>${issue.timeRegistered}</td>
      <td>${issue.system}</td>
      <td class="message-count">${issue.messageCount || 1}</td>
      <td class="subject-cell" title="${escapeHtml(issue.subject)}">${truncate(issue.subject, 40)}</td>
      <td class="description-cell" title="${escapeHtml(issue.description)}">${truncate(issue.description, 60)}</td>
      <td><span class="status-badge status-${getStatusClass(issue.status)}">${issue.status}</span></td>
      <td>${issue.responsible || '-'}</td>
      <td><span class="importance-badge importance-${issue.importance}">${issue.importance}</span></td>
    `
    previewBody.appendChild(row)
  })
}

function showError(error) {
  resultSection.style.display = 'block'
  resultInfo.innerHTML = `
    <div class="error-message">
      <strong>❌ Помилка:</strong><br>
      ${error}
    </div>
  `
}

// ============================================
// Експорт
// ============================================

async function exportToCSV() {
  if (!parsedData || parsedData.length === 0) {
    alert('Немає даних для експорту')
    return
  }

  try {
    exportCsvBtn.disabled = true
    exportCsvBtn.textContent = '⏳ Експорт...'

    const result = await window.electronAPI.exportCSV(parsedData)

    if (result.success) {
      alert(`✅ CSV файл створено!\n\n${result.csvPath}`)
    } else {
      alert('❌ Помилка експорту: ' + result.error)
    }
  } catch (error) {
    console.error('❌ Помилка експорту CSV:', error)
    alert('Помилка: ' + error.message)
  } finally {
    exportCsvBtn.disabled = false
    exportCsvBtn.textContent = '📥 Експорт в CSV'
  }
}

async function exportToJira() {
  if (!parsedData || parsedData.length === 0) {
    alert('Немає даних для експорту')
    return
  }

  if (!confirm(`Створити ${parsedData.length} задач у Jira?`)) {
    return
  }

  try {
    exportJiraBtn.disabled = true
    exportJiraBtn.textContent = '⏳ Створення...'

    const options = {
      host: document.getElementById('jira-host').value,
      email: document.getElementById('jira-email').value,
      apiToken: document.getElementById('jira-token').value,
      projectKey: document.getElementById('jira-project').value,
      issues: parsedData,
    }

    const result = await window.electronAPI.exportToJira(options)

    if (result.success) {
      alert(`✅ Експорт завершено!\n\nСтворено: ${result.created}\nПомилок: ${result.failed}`)
    } else {
      alert('❌ Помилка експорту: ' + result.error)
    }
  } catch (error) {
    console.error('❌ Помилка експорту Jira:', error)
    alert('Помилка: ' + error.message)
  } finally {
    exportJiraBtn.disabled = false
    exportJiraBtn.textContent = '📤 Експорт в Jira'
  }
}

function updateJiraProgress(data) {
  // Можна додати прогрес-бар
  console.log(`Прогрес: ${data.current}/${data.total}`, data.result)
}

// ============================================
// Утиліти
// ============================================

function getStatusClass(status) {
  if (status === 'Вирішено') return 'resolved'
  if (status === 'Вирішено частково') return 'partial'
  if (status === 'У процесі') return 'progress'
  return 'open'
}

function truncate(text, maxLength) {
  if (!text) return ''
  if (text.length <= maxLength) return escapeHtml(text)
  return escapeHtml(text.substring(0, maxLength)) + '...'
}

function escapeHtml(text) {
  if (!text) return ''
  const div = document.createElement('div')
  div.textContent = text
  return div.innerHTML
}

// ============================================
// Запуск
// ============================================

document.addEventListener('DOMContentLoaded', init)
