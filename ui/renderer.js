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
    console.error('electronAPI не знайдено!')
    alert('Помилка: electronAPI не доступний. Перезапустіть додаток.')
    return
  }

  console.log('electronAPI доступний:', Object.keys(window.electronAPI))

  // Завантажуємо конфігурацію
  try {
    const result = await window.electronAPI.loadConfig()

    if (result.success) {
      console.log('Конфігурація завантажена:', result.config)

      // Заповнюємо поля з .env
      document.getElementById('support-emails').value = result.config.supportEmails
      document.getElementById('imap-user').value = result.config.outlookUser
      document.getElementById('imap-host').value = result.config.outlookHost
      document.getElementById('imap-port').value = result.config.outlookPort
      document.getElementById('jira-host').value = result.config.jiraHost
      document.getElementById('jira-email').value = result.config.jiraEmail
      document.getElementById('jira-project').value = result.config.jiraProject
    } else {
      console.warn('Помилка завантаження конфігурації:', result.error)
    }
  } catch (error) {
    console.error('Помилка ініціалізації:', error)
  }

  setupEventListeners()

  // Прогрес Jira
  window.electronAPI.onJiraProgress(data => {
    console.log(`Jira прогрес: ${data.current}/${data.total}`)
    updateJiraProgress(data)
  })
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
  document.getElementById('test-imap-btn').addEventListener('click', testIMAPConnection)
  document.getElementById('select-folders-btn').addEventListener('click', showFolderModal)
  document.getElementById('close-modal').addEventListener('click', closeFolderModal)
  document.getElementById('cancel-folders-btn').addEventListener('click', closeFolderModal)
  document.getElementById('confirm-folders-btn').addEventListener('click', confirmFolderSelection)
  document.getElementById('folder-search').addEventListener('input', filterFolders)

  // Jira
  document.getElementById('connect-jira-btn').addEventListener('click', connectJira)

  // Парсинг
  parseBtn.addEventListener('click', startParsing)

  // Експорт
  exportCsvBtn.addEventListener('click', exportToCSV)
  exportJiraBtn.addEventListener('click', exportToJira)

  // Сортування таблиці
  initTableSorting()

  // Фільтри
  document.getElementById('filter-search').addEventListener('input', applyFilters)
  document.getElementById('filter-system').addEventListener('change', applyFilters)
  document.getElementById('filter-responsible').addEventListener('change', applyFilters)
  document.getElementById('filter-problem-type').addEventListener('change', applyFilters)
  document.getElementById('clear-filters-btn').addEventListener('click', clearFilters)
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
      console.log('PST файл обрано:', result.path)
    }
  } catch (error) {
    console.error('Помилка вибору PST:', error)
    alert('Помилка вибору файлу: ' + error.message)
  }
}

// ============================================
// IMAP
// ============================================

let availableFolders = []
let selectedFolders = []
let imapConnected = false

async function testIMAPConnection() {
  const btn = document.getElementById('test-imap-btn')
  const statusDiv = document.getElementById('imap-status')
  const selectBtn = document.getElementById('select-folders-btn')
  const originalText = btn.textContent

  try {
    btn.disabled = true
    btn.textContent = 'Перевірка...'
    statusDiv.textContent = 'Підключення...'
    statusDiv.className = 'connection-status'

    const credentials = {
      user: document.getElementById('imap-user').value,
      password: document.getElementById('imap-password').value,
      host: document.getElementById('imap-host').value,
      port: parseInt(document.getElementById('imap-port').value),
    }

    if (!credentials.user || !credentials.password) {
      alert('Введіть email та пароль для IMAP')
      statusDiv.textContent = 'Не підключено'
      statusDiv.className = 'connection-status error'
      return
    }

    const result = await window.electronAPI.testIMAPConnection(credentials)

    if (result.success) {
      imapConnected = true
      statusDiv.textContent = '✓ Підключено успішно'
      statusDiv.className = 'connection-status connected'
      selectBtn.disabled = false
      alert('Підключення успішне!\nТепер можна вибрати папки для парсингу.')
    } else {
      imapConnected = false
      statusDiv.textContent = '✗ Помилка підключення'
      statusDiv.className = 'connection-status error'
      alert('Помилка підключення:\n\n' + result.error)
    }
  } catch (error) {
    imapConnected = false
    console.error('Помилка IMAP:', error)
    statusDiv.textContent = '✗ Помилка'
    statusDiv.className = 'connection-status error'
    alert('Помилка: ' + error.message)
  } finally {
    btn.disabled = false
    btn.textContent = originalText
  }
}

async function showFolderModal() {
  const modal = document.getElementById('folder-modal')
  const loadingDiv = document.getElementById('folders-loading')
  const foldersList = document.getElementById('folders-list')

  modal.style.display = 'flex'
  loadingDiv.style.display = 'block'
  foldersList.innerHTML = ''

  try {
    const credentials = {
      user: document.getElementById('imap-user').value,
      password: document.getElementById('imap-password').value,
      host: document.getElementById('imap-host').value,
      port: parseInt(document.getElementById('imap-port').value),
    }

    const result = await window.electronAPI.getIMAPFolders(credentials)

    if (result.success) {
      availableFolders = result.folders
      renderFolders(availableFolders)
    } else {
      alert('Помилка отримання папок:\n\n' + result.error)
      closeFolderModal()
    }
  } catch (error) {
    console.error('Помилка:', error)
    alert('Помилка: ' + error.message)
    closeFolderModal()
  } finally {
    loadingDiv.style.display = 'none'
  }
}

function renderFolders(folders) {
  const foldersList = document.getElementById('folders-list')
  foldersList.innerHTML = ''

  folders.forEach(folder => {
    const item = document.createElement('div')
    item.className = 'folder-item'
    if (selectedFolders.includes(folder.name)) {
      item.classList.add('selected')
    }

    const checkbox = document.createElement('input')
    checkbox.type = 'checkbox'
    checkbox.checked = selectedFolders.includes(folder.name)
    checkbox.addEventListener('change', () => toggleFolderSelection(folder.name, checkbox.checked))

    const info = document.createElement('div')
    info.className = 'folder-info'

    const name = document.createElement('div')
    name.className = 'folder-name'
    name.textContent = folder.name

    const stats = document.createElement('div')
    stats.className = 'folder-stats'
    stats.textContent = `${folder.total || 0} листів (${folder.unseen || 0} непрочитаних)`

    info.appendChild(name)
    info.appendChild(stats)

    item.appendChild(checkbox)
    item.appendChild(info)

    item.addEventListener('click', e => {
      if (e.target !== checkbox) {
        checkbox.checked = !checkbox.checked
        toggleFolderSelection(folder.name, checkbox.checked)
      }
    })

    foldersList.appendChild(item)
  })
}

function toggleFolderSelection(folderName, selected) {
  if (selected) {
    if (!selectedFolders.includes(folderName)) {
      selectedFolders.push(folderName)
    }
  } else {
    selectedFolders = selectedFolders.filter(f => f !== folderName)
  }

  // Update visual state
  const items = document.querySelectorAll('.folder-item')
  items.forEach(item => {
    const checkbox = item.querySelector('input[type="checkbox"]')
    const folderNameEl = item.querySelector('.folder-name')
    if (folderNameEl && folderNameEl.textContent === folderName) {
      if (selected) {
        item.classList.add('selected')
      } else {
        item.classList.remove('selected')
      }
    }
  })
}

function filterFolders() {
  const searchInput = document.getElementById('folder-search')
  const searchText = searchInput.value.toLowerCase()

  const filtered = availableFolders.filter(folder => folder.name.toLowerCase().includes(searchText))

  renderFolders(filtered)
}

function closeFolderModal() {
  const modal = document.getElementById('folder-modal')
  modal.style.display = 'none'
}

function confirmFolderSelection() {
  if (selectedFolders.length === 0) {
    alert('Виберіть хоча б одну папку')
    return
  }

  // Update UI with selected folders
  const display = document.getElementById('selected-folders-display')
  const list = document.getElementById('selected-folders-list')

  list.innerHTML = ''
  selectedFolders.forEach(folder => {
    const li = document.createElement('li')
    li.textContent = folder
    list.appendChild(li)
  })

  display.style.display = 'block'
  parseBtn.disabled = false

  closeFolderModal()
  alert(`Вибрано ${selectedFolders.length} папок для парсингу`)
}

// ============================================
// Jira
// ============================================

async function connectJira() {
  const btn = document.getElementById('connect-jira-btn')
  const originalText = btn.textContent

  try {
    btn.disabled = true
    btn.textContent = 'Підключення...'

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
      console.log('Jira підключено:', result)
      parseBtn.disabled = false
      alert(`Успішно підключено до Jira!\n\nКористувач: ${result.user.displayName}\nПроект: ${result.project}`)
    } else {
      alert('Помилка підключення:\n\n' + result.error)
    }
  } catch (error) {
    console.error('Помилка Jira:', error)
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
    console.log('DEBUG: Початок парсингу, джерело:', selectedSource)
    loading.style.display = 'block'
    resultSection.style.display = 'none'
    previewSection.style.display = 'none'
    parseBtn.disabled = true

    let result

    if (selectedSource === 'pst') {
      console.log('DEBUG: Викликаємо parsePST()')
      result = await parsePST()
    } else if (selectedSource === 'imap') {
      console.log('DEBUG: Викликаємо parseIMAP()')
      result = await parseIMAP()
    } else if (selectedSource === 'jira') {
      console.log('DEBUG: Викликаємо parseJira()')
      result = await parseJira()
    }

    console.log('DEBUG: Отримано результат від backend:')
    console.log('  - Success:', result.success)
    console.log('  - Data length:', result.data ? result.data.length : 0)
    console.log('  - Stats:', result.stats)
    console.log('  - Error:', result.error)

    loading.style.display = 'none'
    parseBtn.disabled = false

    if (result.success) {
      console.log('DEBUG: Результат успішний, зберігаємо parsedData')
      parsedData = result.data
      console.log('DEBUG: parsedData збережено, кількість:', parsedData.length)
      displayResults(result)
    } else {
      console.error('DEBUG: Результат містить помилку:', result.error)
      showError(result.error)
    }
  } catch (error) {
    loading.style.display = 'none'
    parseBtn.disabled = false
    console.error('Помилка парсингу:', error)
    console.error('DEBUG: Stack trace:', error.stack)
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

  console.log('DEBUG: Опції для PST парсингу:', {
    pstPath: options.pstPath,
    supportEmails: options.supportEmails,
    keywords: options.keywords,
    startDate: options.startDate,
    endDate: options.endDate,
    batchSize: options.batchSize,
  })

  if (!options.pstPath) {
    throw new Error('Оберіть PST файл')
  }

  console.log('DEBUG: Викликаємо electronAPI.parsePST...')
  const result = await window.electronAPI.parsePST(options)
  console.log('DEBUG: Отримано відповідь від parsePST:', result)

  return result
}

async function parseIMAP() {
  if (selectedFolders.length === 0) {
    throw new Error('Виберіть хоча б одну папку для парсингу')
  }

  const options = {
    user: document.getElementById('imap-user').value,
    password: document.getElementById('imap-password').value,
    host: document.getElementById('imap-host').value,
    port: parseInt(document.getElementById('imap-port').value),
    folders: selectedFolders,
    supportEmails: document.getElementById('support-emails').value,
    keywords: document.getElementById('keywords').value,
    startDate: document.getElementById('start-date').value,
    endDate: document.getElementById('end-date').value,
  }

  if (!options.user || !options.password) {
    throw new Error('Введіть credentials для IMAP')
  }

  console.log(`Парсинг з IMAP: ${selectedFolders.length} папок`)
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
  console.log('DEBUG: displayResults викликано з даними:', {
    dataLength: result.data ? result.data.length : 0,
    stats: result.stats,
  })

  const stats = result.stats

  console.log('DEBUG: Статистика для відображення:', stats)

  resultSection.style.display = 'block'
  resultInfo.innerHTML = `
    <div class="success-message">
      <strong>Парсинг завершено!</strong><br><br>
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
  console.log('DEBUG: Викликаємо displayTable з', result.data.length, 'записами')

  // Заповнюємо опції фільтрів
  populateFilterOptions()

  displayTable(result.data)
}

function displayTable(issues) {
  console.log('DEBUG: displayTable викликано, issues:', issues.length)
  console.log('DEBUG: Перше issue:', issues[0])

  previewBody.innerHTML = ''

  issues.forEach((issue, index) => {
    if (index < 3) {
      console.log(`DEBUG: Issue #${index}:`, issue)
    }

    const row = document.createElement('tr')
    row.innerHTML = `
      <td>${issue.dateRegistered}</td>
      <td>${issue.timeRegistered}</td>
      <td>${issue.system}</td>
      <td class="problem-type">${issue.problemType || '-'}</td>
      <td class="thread-position">${issue.threadPosition || '-'}</td>
      <td class="subject-cell" title="${escapeHtml(issue.subject)}">${truncate(issue.subject, 30)}</td>
      <td class="description-cell" title="${escapeHtml(issue.description)}">${truncate(issue.description, 50)}</td>
      <td class="description-cell" title="${escapeHtml(issue.solution)}">${truncate(issue.solution, 50)}</td>
      <td><span class="status-badge status-${getStatusClass(issue.status)}">${issue.status}</span></td>
      <td>${issue.responsible || '-'}</td>
      <td><span class="importance-badge importance-${issue.importance}">${issue.importance}</span></td>
    `
    previewBody.appendChild(row)
  })

  console.log('DEBUG: Таблиця відображена, рядків:', issues.length)
}

function showError(error) {
  resultSection.style.display = 'block'
  resultInfo.innerHTML = `
    <div class="error-message">
      <strong>Помилка:</strong><br>
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
    exportCsvBtn.textContent = 'Експорт...'

    const result = await window.electronAPI.exportCSV(parsedData)

    if (result.success) {
      alert(`CSV файл створено!\n\n${result.csvPath}`)
    } else {
      alert('Помилка експорту: ' + result.error)
    }
  } catch (error) {
    console.error('Помилка експорту CSV:', error)
    alert('Помилка: ' + error.message)
  } finally {
    exportCsvBtn.disabled = false
    exportCsvBtn.textContent = 'Експорт в CSV'
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
    exportJiraBtn.textContent = 'Створення...'

    const options = {
      host: document.getElementById('jira-host').value,
      email: document.getElementById('jira-email').value,
      apiToken: document.getElementById('jira-token').value,
      projectKey: document.getElementById('jira-project').value,
      issues: parsedData,
    }

    const result = await window.electronAPI.exportToJira(options)

    if (result.success) {
      alert(`Експорт завершено!\n\nСтворено: ${result.created}\nПомилок: ${result.failed}`)
    } else {
      alert('Помилка експорту: ' + result.error)
    }
  } catch (error) {
    console.error('Помилка експорту Jira:', error)
    alert('Помилка: ' + error.message)
  } finally {
    exportJiraBtn.disabled = false
    exportJiraBtn.textContent = 'Експорт в Jira'
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
// Фільтрація
// ============================================

function populateFilterOptions() {
  if (!parsedData || parsedData.length === 0) return

  // Унікальні системи
  const systems = [...new Set(parsedData.map(item => item.system).filter(Boolean))].sort()
  const systemSelect = document.getElementById('filter-system')
  systemSelect.innerHTML = '<option value="">Всі системи</option>'
  systems.forEach(system => {
    const option = document.createElement('option')
    option.value = system
    option.textContent = system
    systemSelect.appendChild(option)
  })

  // Унікальні відповідальні
  const responsible = [...new Set(parsedData.map(item => item.responsible).filter(Boolean))].sort()
  const responsibleSelect = document.getElementById('filter-responsible')
  responsibleSelect.innerHTML = '<option value="">Всі відповідальні</option>'
  responsible.forEach(person => {
    const option = document.createElement('option')
    option.value = person
    option.textContent = person
    responsibleSelect.appendChild(option)
  })

  // Унікальні типи проблем
  const problemTypes = [...new Set(parsedData.map(item => item.problemType).filter(Boolean))].sort()
  const problemTypeSelect = document.getElementById('filter-problem-type')
  problemTypeSelect.innerHTML = '<option value="">Всі типи</option>'
  problemTypes.forEach(type => {
    const option = document.createElement('option')
    option.value = type
    option.textContent = type
    problemTypeSelect.appendChild(option)
  })
}

function applyFilters() {
  if (!parsedData || parsedData.length === 0) return

  const searchText = document.getElementById('filter-search').value.toLowerCase()
  const systemFilter = document.getElementById('filter-system').value
  const responsibleFilter = document.getElementById('filter-responsible').value
  const problemTypeFilter = document.getElementById('filter-problem-type').value

  let filtered = parsedData.filter(issue => {
    // Пошук по тексту
    if (searchText) {
      const searchableText = `${issue.subject} ${issue.description} ${issue.requestText} ${issue.responseText}`.toLowerCase()
      if (!searchableText.includes(searchText)) return false
    }

    // Фільтр по системі
    if (systemFilter && issue.system !== systemFilter) return false

    // Фільтр по відповідальному
    if (responsibleFilter && issue.responsible !== responsibleFilter) return false

    // Фільтр по типу проблеми
    if (problemTypeFilter && issue.problemType !== problemTypeFilter) return false

    return true
  })

  // Застосовуємо сортування якщо є
  if (currentSort.column) {
    filtered = filtered.sort((a, b) => compareValues(a, b, currentSort.column, currentSort.direction))
  }

  displayTable(filtered)

  // Оновлюємо лічильник
  document.querySelector('.preview-header h2').textContent = `Звернення (${filtered.length} з ${parsedData.length})`
}

function clearFilters() {
  document.getElementById('filter-search').value = ''
  document.getElementById('filter-system').value = ''
  document.getElementById('filter-responsible').value = ''
  document.getElementById('filter-problem-type').value = ''

  // Скидаємо сортування
  currentSort.column = null
  currentSort.direction = null
  updateSortIndicators()

  displayTable(parsedData)
  document.querySelector('.preview-header h2').textContent = 'Звернення'
}

// ============================================
// Сортування таблиці
// ============================================

let currentSort = {
  column: null,
  direction: null, // 'asc' or 'desc'
}

function initTableSorting() {
  const sortableHeaders = document.querySelectorAll('.sortable')

  sortableHeaders.forEach(header => {
    header.addEventListener('click', () => {
      const column = header.dataset.column
      sortTable(column)
    })
  })
}

function sortTable(column) {
  if (!parsedData || parsedData.length === 0) return

  // Визначаємо напрямок сортування
  if (currentSort.column === column) {
    // Переключаємо напрямок
    if (currentSort.direction === 'asc') {
      currentSort.direction = 'desc'
    } else if (currentSort.direction === 'desc') {
      // Скидаємо сортування
      currentSort.column = null
      currentSort.direction = null
    } else {
      currentSort.direction = 'asc'
    }
  } else {
    // Нова колонка - сортуємо за зростанням
    currentSort.column = column
    currentSort.direction = 'asc'
  }

  // Оновлюємо візуальні індикатори
  updateSortIndicators()

  // Якщо скинуто сортування - повертаємо оригінальний порядок
  if (!currentSort.column) {
    displayTable(parsedData)
    return
  }

  // Сортуємо дані
  const sortedData = [...parsedData].sort((a, b) => {
    return compareValues(a, b, column, currentSort.direction)
  })

  // Відображаємо відсортовані дані
  displayTable(sortedData)
}

function compareValues(a, b, column, direction) {
  let aValue = a[column]
  let bValue = b[column]

  // Обробка порожніх значень
  if (aValue === null || aValue === undefined) aValue = ''
  if (bValue === null || bValue === undefined) bValue = ''

  // Спеціальна обробка для дат
  if (column === 'dateRegistered') {
    aValue = new Date(aValue + ' ' + (a.timeRegistered || '00:00:00'))
    bValue = new Date(bValue + ' ' + (b.timeRegistered || '00:00:00'))
  }

  // Спеціальна обробка для важливості
  if (column === 'importance') {
    const importanceOrder = { Високий: 3, Середній: 2, Низький: 1 }
    aValue = importanceOrder[aValue] || 0
    bValue = importanceOrder[bValue] || 0
  }

  // Порівняння
  let comparison = 0

  if (aValue > bValue) {
    comparison = 1
  } else if (aValue < bValue) {
    comparison = -1
  }

  // Застосовуємо напрямок
  return direction === 'desc' ? comparison * -1 : comparison
}

function updateSortIndicators() {
  // Видаляємо всі індикатори
  document.querySelectorAll('.sortable').forEach(header => {
    header.classList.remove('sort-asc', 'sort-desc')
  })

  // Додаємо індикатор до активної колонки
  if (currentSort.column) {
    const activeHeader = document.querySelector(`[data-column="${currentSort.column}"]`)
    if (activeHeader) {
      activeHeader.classList.add(`sort-${currentSort.direction}`)
    }
  }
}

// ============================================
// Запуск
// ============================================

document.addEventListener('DOMContentLoaded', init)
