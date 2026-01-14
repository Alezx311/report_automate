// Глобальні змінні
let parsedData = null
let selectedSource = 'pst' // 'pst', 'graph', 'jira'

// Елементи UI
const sourceButtons = {
  pst: document.getElementById('source-pst'),
  graph: document.getElementById('source-graph'),
  jira: document.getElementById('source-jira'),
}

const sourcePanels = {
  pst: document.getElementById('pst-panel'),
  graph: document.getElementById('graph-panel'),
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
      document.getElementById('jira-host').value = result.config.jiraHost
      document.getElementById('jira-email').value = result.config.jiraEmail
      document.getElementById('jira-project').value = result.config.jiraProject
      // Graph API
      document.getElementById('graph-tenant').value = result.config.azureTenant
      document.getElementById('graph-client-id').value = result.config.azureClientId
      document.getElementById('graph-user').value = result.config.graphUser
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

  // Graph API
  document.getElementById('test-graph-btn').addEventListener('click', testGraphConnection)
  document.getElementById('select-graph-folders-btn').addEventListener('click', showGraphFolderModal)
  document.getElementById('close-modal').addEventListener('click', closeFolderModal)
  document.getElementById('cancel-folders-btn').addEventListener('click', closeFolderModal)
  document.getElementById('confirm-folders-btn').addEventListener('click', confirmFolderSelection)
  document.getElementById('folder-search').addEventListener('input', filterGraphFolders)

  // Jira
  document.getElementById('connect-jira-btn').addEventListener('click', connectJira)

  // Cache
  document.getElementById('manage-cache-btn').addEventListener('click', showCacheModal)
  document.getElementById('close-cache-modal').addEventListener('click', closeCacheModal)
  document.getElementById('close-cache-modal-btn').addEventListener('click', closeCacheModal)
  document.getElementById('refresh-cache-list-btn').addEventListener('click', loadCacheList)
  document.getElementById('clear-all-cache-btn').addEventListener('click', clearAllCache)

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
// Graph API
// ============================================

let availableGraphFolders = []
let selectedGraphFolders = []
let graphConnected = false

async function testGraphConnection() {
  const btn = document.getElementById('test-graph-btn')
  const statusDiv = document.getElementById('graph-status')
  const selectBtn = document.getElementById('select-graph-folders-btn')
  const originalText = btn.textContent

  try {
    btn.disabled = true
    btn.textContent = 'Перевірка...'
    statusDiv.textContent = 'Підключення...'
    statusDiv.className = 'connection-status'

    const credentials = {
      tenant: document.getElementById('graph-tenant').value,
      clientId: document.getElementById('graph-client-id').value,
      clientSecret: document.getElementById('graph-client-secret').value,
      user: document.getElementById('graph-user').value,
      password: document.getElementById('graph-password').value,
    }

    if (!credentials.tenant || !credentials.clientId || !credentials.clientSecret || !credentials.user || !credentials.password) {
      alert('Заповніть всі поля для Graph API')
      statusDiv.textContent = 'Не підключено'
      statusDiv.className = 'connection-status error'
      return
    }

    const result = await window.electronAPI.testGraphConnection(credentials)

    if (result.success) {
      graphConnected = true
      statusDiv.textContent = `✓ Підключено як ${result.user}`
      statusDiv.className = 'connection-status connected'
      selectBtn.disabled = false
      alert(`Підключення успішне!\nКористувач: ${result.displayName || result.user}\nТепер можна вибрати папки для парсингу.`)
    } else {
      graphConnected = false
      statusDiv.textContent = '✗ Помилка підключення'
      statusDiv.className = 'connection-status error'
      alert('Помилка підключення:\n\n' + result.error)
    }
  } catch (error) {
    graphConnected = false
    console.error('Помилка Graph API:', error)
    statusDiv.textContent = '✗ Помилка'
    statusDiv.className = 'connection-status error'
    alert('Помилка: ' + error.message)
  } finally {
    btn.disabled = false
    btn.textContent = originalText
  }
}

async function showGraphFolderModal() {
  const modal = document.getElementById('folder-modal')
  const loadingDiv = document.getElementById('folders-loading')
  const foldersList = document.getElementById('folders-list')

  modal.style.display = 'flex'
  loadingDiv.style.display = 'block'
  foldersList.innerHTML = ''

  try {
    const credentials = {
      tenant: document.getElementById('graph-tenant').value,
      clientId: document.getElementById('graph-client-id').value,
      clientSecret: document.getElementById('graph-client-secret').value,
      user: document.getElementById('graph-user').value,
      password: document.getElementById('graph-password').value,
    }

    const result = await window.electronAPI.getGraphFolders(credentials)

    if (result.success) {
      availableGraphFolders = result.folders
      renderGraphFolders(availableGraphFolders)
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

function renderGraphFolders(folders) {
  const foldersList = document.getElementById('folders-list')
  foldersList.innerHTML = ''

  folders.forEach(folder => {
    const item = document.createElement('div')
    item.className = 'folder-item'
    if (selectedGraphFolders.includes(folder.name)) {
      item.classList.add('selected')
    }

    const checkbox = document.createElement('input')
    checkbox.type = 'checkbox'
    checkbox.checked = selectedGraphFolders.includes(folder.name)
    checkbox.addEventListener('change', () => toggleGraphFolderSelection(folder.name, checkbox.checked))

    const info = document.createElement('div')
    info.className = 'folder-info'

    const name = document.createElement('div')
    name.className = 'folder-name'
    name.textContent = folder.name

    const stats = document.createElement('div')
    stats.className = 'folder-stats'
    stats.textContent = `${folder.totalItemCount || 0} листів (${folder.unreadItemCount || 0} непрочитаних)`

    info.appendChild(name)
    info.appendChild(stats)

    item.appendChild(checkbox)
    item.appendChild(info)

    item.addEventListener('click', e => {
      if (e.target !== checkbox) {
        checkbox.checked = !checkbox.checked
        toggleGraphFolderSelection(folder.name, checkbox.checked)
      }
    })

    foldersList.appendChild(item)
  })
}

function toggleGraphFolderSelection(folderName, isSelected) {
  if (isSelected) {
    if (!selectedGraphFolders.includes(folderName)) {
      selectedGraphFolders.push(folderName)
    }
  } else {
    selectedGraphFolders = selectedGraphFolders.filter(f => f !== folderName)
  }

  // Update visual state
  const items = document.querySelectorAll('.folder-item')
  items.forEach(item => {
    const folderNameEl = item.querySelector('.folder-name')
    if (folderNameEl && folderNameEl.textContent === folderName) {
      if (isSelected) {
        item.classList.add('selected')
      } else {
        item.classList.remove('selected')
      }
    }
  })
}

function filterGraphFolders() {
  const searchInput = document.getElementById('folder-search')
  const searchText = searchInput.value.toLowerCase()

  const filtered = availableGraphFolders.filter(folder => folder.name.toLowerCase().includes(searchText))

  renderGraphFolders(filtered)
}

function closeFolderModal() {
  const modal = document.getElementById('folder-modal')
  modal.style.display = 'none'
}

function confirmFolderSelection() {
  if (selectedGraphFolders.length === 0) {
    alert('Виберіть хоча б одну папку')
    return
  }

  // Update UI with selected folders
  const display = document.getElementById('selected-graph-folders-display')
  const list = document.getElementById('selected-graph-folders-list')

  list.innerHTML = ''
  selectedGraphFolders.forEach(folder => {
    const li = document.createElement('li')
    li.textContent = folder
    list.appendChild(li)
  })

  display.style.display = 'block'
  parseBtn.disabled = false

  closeFolderModal()
  alert(`Вибрано ${selectedGraphFolders.length} папок для парсингу`)
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
    } else if (selectedSource === 'graph') {
      console.log('DEBUG: Викликаємо parseGraph()')
      result = await parseGraph()
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
    useAggressiveClean: document.getElementById('aggressive-clean').checked,
    batchSize: 100,
  }

  console.log('DEBUG: Опції для PST парсингу:', {
    pstPath: options.pstPath,
    supportEmails: options.supportEmails,
    keywords: options.keywords,
    startDate: options.startDate,
    endDate: options.endDate,
    useAggressiveClean: options.useAggressiveClean,
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

async function parseGraph() {
  if (selectedGraphFolders.length === 0) {
    throw new Error('Виберіть хоча б одну папку для парсингу')
  }

  const options = {
    tenant: document.getElementById('graph-tenant').value,
    clientId: document.getElementById('graph-client-id').value,
    clientSecret: document.getElementById('graph-client-secret').value,
    user: document.getElementById('graph-user').value,
    password: document.getElementById('graph-password').value,
    folders: selectedGraphFolders,
    supportEmails: document.getElementById('support-emails').value,
    keywords: document.getElementById('keywords').value,
    startDate: document.getElementById('start-date').value,
    endDate: document.getElementById('end-date').value,
  }

  if (!options.tenant || !options.clientId || !options.clientSecret || !options.user || !options.password) {
    throw new Error('Заповніть всі поля для Graph API')
  }

  console.log(`Парсинг з Graph API: ${selectedGraphFolders.length} папок`)
  const result = await window.electronAPI.parseGraph(options)

  // Перевіряємо чи потрібно зберегти в кеш
  if (result.success && document.getElementById('save-to-cache-checkbox').checked) {
    try {
      console.log('Збереження в кеш...')
      const cacheResult = await window.electronAPI.saveToCache({
        source: 'graph-api',
        data: result.data,
        startDate: options.startDate,
        endDate: options.endDate,
        folders: options.folders,
        supportEmails: options.supportEmails,
        keywords: options.keywords,
      })

      if (cacheResult.success) {
        console.log(`✅ Збережено ${cacheResult.messageCount} листів в кеш: ${cacheResult.fileName}`)
        alert(`✅ Дані збережено в кеш!\nФайл: ${cacheResult.fileName}\nЛистів: ${cacheResult.messageCount}`)
      }
    } catch (error) {
      console.error('Помилка збереження в кеш:', error)
      // Не переривуємо процес, просто логуємо помилку
    }
  }

  return result
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
// Cache Management
// ============================================

async function showCacheModal() {
  const modal = document.getElementById('cache-modal')
  modal.style.display = 'flex'

  await loadCacheStats()
  await loadCacheList()
}

function closeCacheModal() {
  const modal = document.getElementById('cache-modal')
  modal.style.display = 'none'
}

async function loadCacheStats() {
  try {
    const result = await window.electronAPI.getCacheStats()

    if (result.success && result.stats) {
      const stats = result.stats
      document.getElementById('cache-stats-content').innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px;">
          <div>
            <div style="font-size: 12px; color: #6b7280;">Файлів кешу</div>
            <div style="font-size: 20px; font-weight: 600; color: #667eea;">${stats.totalFiles}</div>
          </div>
          <div>
            <div style="font-size: 12px; color: #6b7280;">Загальний розмір</div>
            <div style="font-size: 20px; font-weight: 600; color: #667eea;">${stats.totalSizeFormatted}</div>
          </div>
          <div>
            <div style="font-size: 12px; color: #6b7280;">Всього листів</div>
            <div style="font-size: 20px; font-weight: 600; color: #667eea;">${stats.totalMessages}</div>
          </div>
        </div>
      `
    }
  } catch (error) {
    console.error('Failed to load cache stats:', error)
  }
}

async function loadCacheList() {
  const loadingDiv = document.getElementById('cache-loading')
  const listDiv = document.getElementById('cache-files-list')

  loadingDiv.style.display = 'block'
  listDiv.innerHTML = ''

  try {
    const result = await window.electronAPI.listCacheFiles()

    if (result.success && result.files.length > 0) {
      listDiv.innerHTML = result.files.map(file => `
        <div class="cache-file-item">
          <div class="cache-file-info">
            <div class="cache-file-name">${file.source.toUpperCase()} - ${formatCacheDate(file.cachedAt)}</div>
            <div class="cache-file-meta">
              <span>📅 ${formatDateRange(file.dateRange)}</span>
              <span>📊 ${file.messageCount} листів</span>
              <span>💾 ${file.sizeFormatted}</span>
              ${file.folders.length > 0 ? `<span>📁 ${file.folders.slice(0, 2).join(', ')}${file.folders.length > 2 ? '...' : ''}</span>` : ''}
            </div>
          </div>
          <div class="cache-file-actions">
            <button class="btn btn-success btn-small" onclick="loadCachedData('${file.fileName}')">📂 Завантажити</button>
            <button class="btn btn-danger btn-small" onclick="deleteCachedFile('${file.fileName}')">🗑️</button>
          </div>
        </div>
      `).join('')
    } else {
      listDiv.innerHTML = `
        <div class="cache-empty">
          <p>📭 Кеш порожній</p>
          <p>Завантажте дані з Graph API і відмітьте "Зберегти в кеш"</p>
        </div>
      `
    }

    await loadCacheStats()
  } catch (error) {
    console.error('Failed to load cache list:', error)
    listDiv.innerHTML = '<div class="cache-empty"><p>❌ Помилка завантаження кешу</p></div>'
  } finally {
    loadingDiv.style.display = 'none'
  }
}

function formatCacheDate(dateStr) {
  const date = new Date(dateStr)
  return date.toLocaleString('uk-UA', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  })
}

function formatDateRange(range) {
  if (!range || (!range.start && !range.end)) return 'Всі дати'

  const start = range.start ? new Date(range.start).toLocaleDateString('uk-UA') : '...'
  const end = range.end ? new Date(range.end).toLocaleDateString('uk-UA') : '...'

  return `${start} - ${end}`
}

async function loadCachedData(fileName) {
  try {
    loading.style.display = 'block'
    closeCacheModal()

    const result = await window.electronAPI.loadFromCache(fileName)

    if (result.success) {
      // Генеруємо звіт з кешованих даних
      const reportGenerator = {
        processMessages: (messages) => {
          // Тут використовуємо той самий код, що і для Graph API
          return {
            issues: messages,
            stats: result.stats || {}
          }
        }
      }

      parsedData = result.data
      displayResults({
        success: true,
        data: result.data,
        stats: result.stats || {
          totalThreads: 0,
          total: result.data.length,
          resolved: 0,
          inProgress: result.data.length,
          avgMessagesPerIssue: 0
        }
      })

      alert(`✅ Завантажено ${result.data.length} листів з кешу`)
    } else {
      alert('❌ Помилка завантаження з кешу:\n\n' + result.error)
    }
  } catch (error) {
    console.error('Cache load error:', error)
    alert('❌ Помилка: ' + error.message)
  } finally {
    loading.style.display = 'none'
  }
}

async function deleteCachedFile(fileName) {
  if (!confirm(`Видалити файл кешу?\n\n${fileName}`)) {
    return
  }

  try {
    const result = await window.electronAPI.deleteCacheFile(fileName)

    if (result.success) {
      await loadCacheList()
    } else {
      alert('❌ Помилка видалення: ' + result.error)
    }
  } catch (error) {
    console.error('Delete cache error:', error)
    alert('❌ Помилка: ' + error.message)
  }
}

async function clearAllCache() {
  if (!confirm('Видалити весь кеш?\n\nЦю дію не можна скасувати.')) {
    return
  }

  try {
    const result = await window.electronAPI.clearAllCache()

    if (result.success) {
      alert(`✅ Видалено ${result.deletedCount} файлів`)
      await loadCacheList()
    } else {
      alert('❌ Помилка: ' + result.error)
    }
  } catch (error) {
    console.error('Clear cache error:', error)
    alert('❌ Помилка: ' + error.message)
  }
}

// ============================================
// Запуск
// ============================================

document.addEventListener('DOMContentLoaded', init)
