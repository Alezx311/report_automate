let selectedPSTPath = null
let parsedData = null

// Елементи UI
const selectFileBtn = document.getElementById('select-file-btn')
const pstPathInput = document.getElementById('pst-path')
const parseBtn = document.getElementById('parse-btn')
const exportCsvBtn = document.getElementById('export-csv-btn')
const supportEmailsInput = document.getElementById('support-emails')
const keywordsInput = document.getElementById('keywords')
const filterTextInput = document.getElementById('filter-text')
const resultSection = document.getElementById('result-section')
const resultInfo = document.getElementById('result-info')
const previewSection = document.getElementById('preview-section')
const previewBody = document.getElementById('preview-body')
const loading = document.getElementById('loading')

// Вибір PST файлу
selectFileBtn.addEventListener('click', async () => {
  const filePath = await window.electronAPI.selectPSTFile()

  if (filePath) {
    selectedPSTPath = filePath
    pstPathInput.value = filePath
    parseBtn.disabled = false
  }
})

// Парсинг PST
parseBtn.addEventListener('click', async () => {
  if (!selectedPSTPath) {
    alert('Будь ласка, оберіть PST файл')
    return
  }

  const supportEmails = supportEmailsInput.value.trim()
  const keywords = keywordsInput.value.trim()
  const filterText = filterTextInput.value.trim()
  const startDate = document.getElementById('startDate').value
  const endDate = document.getElementById('endDate').value
  const batchSize = parseInt(document.getElementById('batchSize').value) || 100
  const ignoreInvalidDates = document.getElementById('ignoreInvalidDates').checked

  if (!supportEmails) {
    alert('Будь ласка, вкажіть email адреси техпідтримки')
    return
  }

  // Показати завантаження
  loading.style.display = 'block'
  resultSection.style.display = 'none'
  previewSection.style.display = 'none'
  parseBtn.disabled = true

  try {
    const result = await window.electronAPI.parsePST({
      pstPath: selectedPSTPath,
      supportEmails,
      keywords,
      filterText,
    })

    loading.style.display = 'none'
    parseBtn.disabled = false

    if (result.success) {
      parsedData = result.data

      // Показати результат зі статистикою
      resultSection.style.display = 'block'
      resultInfo.innerHTML = `
        <div class="success-message">
          <strong>✅ Парсинг завершено!</strong><br><br>
          <div class="stats-grid">
            <div class="stat-item">
              <div class="stat-label">Всього threads:</div>
              <div class="stat-value">${result.stats.totalThreads}</div>
            </div>
            <div class="stat-item">
              <div class="stat-label">Всього звернень:</div>
              <div class="stat-value">${result.stats.total}</div>
            </div>
            <div class="stat-item stat-resolved">
              <div class="stat-label">Вирішено:</div>
              <div class="stat-value">${result.stats.resolved}</div>
            </div>
            <div class="stat-item stat-progress">
              <div class="stat-label">У процесі:</div>
              <div class="stat-value">${result.stats.inProgress}</div>
            </div>
          </div>
          <p style="margin-top: 15px;">Відредагуйте дані нижче та натисніть "Експортувати в CSV"</p>
        </div>
      `

      // Показати таблицю для редагування
      previewSection.style.display = 'block'
      displayEditableTable(parsedData)
    } else {
      // Показати помилку
      resultSection.style.display = 'block'
      resultInfo.innerHTML = `
        <div class="error-message">
          <strong>❌ Помилка:</strong><br>
          ${result.error}
        </div>
      `
    }
  } catch (error) {
    loading.style.display = 'none'
    parseBtn.disabled = false
    resultSection.style.display = 'block'
    resultInfo.innerHTML = `
      <div class="error-message">
        <strong>❌ Критична помилка:</strong><br>
        ${error.message}
      </div>
    `
  }
})

// Експорт в CSV
exportCsvBtn.addEventListener('click', async () => {
  if (!parsedData) {
    alert('Немає даних для експорту')
    return
  }

  // Збираємо відредаговані дані з таблиці
  const updatedData = []
  const rows = previewBody.querySelectorAll('tr')

  rows.forEach((row, index) => {
    const issue = parsedData[index]

    // Беремо textarea якщо є, інакше оригінальний опис
    const descriptionTextarea = row.querySelector('.edit-description')
    const description = descriptionTextarea ? descriptionTextarea.value : issue.description

    updatedData.push({
      dateRegistered: row.querySelector('.edit-date').value,
      timeRegistered: row.querySelector('.edit-time').value,
      system: row.querySelector('.edit-system').value,
      ticketId: '',
      communication: 'Пошта',
      description: description,
      status: row.querySelector('.edit-status').value,
      responsible: row.querySelector('.edit-responsible').value,
      solution: issue.solution,
      dateResolved: issue.dateResolved,
      timeResolved: issue.timeResolved,
      importance: row.querySelector('.edit-importance').value,
    })
  })

  exportCsvBtn.disabled = true
  exportCsvBtn.textContent = '⏳ Експорт...'

  try {
    const result = await window.electronAPI.exportCSV(updatedData)

    if (result.success) {
      alert(`✅ CSV файл створено!\n\nШлях: ${result.csvPath}`)
      resultInfo.innerHTML = `
        <div class="success-message">
          <strong>✅ CSV експортовано!</strong><br>
          Файл створено: <code>${result.csvPath}</code>
        </div>
      `
    } else {
      alert(`❌ Помилка експорту: ${result.error}`)
    }
  } catch (error) {
    alert(`❌ Критична помилка: ${error.message}`)
  } finally {
    exportCsvBtn.disabled = false
    exportCsvBtn.textContent = '📥 Експортувати в CSV'
  }
})

function displayEditableTable(issues) {
  previewBody.innerHTML = ''

  issues.forEach((issue, index) => {
    const row = document.createElement('tr')

    // Скорочений опис для відображення (перші 200 символів)
    const shortDescription = issue.description.substring(0, 200)

    row.innerHTML = `
      <td><input type="date" class="edit-date table-input" value="${issue.dateRegistered}"></td>
      <td><input type="time" class="edit-time table-input" value="${issue.timeRegistered}"></td>
      <td><input type="text" class="edit-system table-input" value="${issue.system}"></td>
      <td class="subject-cell" title="${escapeHtml(issue.subject)}">${truncate(issue.subject, 50)}</td>
      <td>
        <div class="description-preview" title="Натисніть для перегляду повного опису">
          ${truncate(shortDescription, 100)}
        </div>
        <button class="btn-small btn-view" onclick="viewFullDescription(${index})">👁️ Переглянути</button>
        <textarea class="edit-description table-textarea" rows="3" style="display:none;">${escapeHtml(
          issue.description,
        )}</textarea>
      </td>
      <td>
        <select class="edit-status table-select">
          <option value="Вирішено" ${issue.status === 'Вирішено' ? 'selected' : ''}>Вирішено</option>
          <option value="У процесі" ${issue.status === 'У процесі' ? 'selected' : ''}>У процесі</option>
        </select>
      </td>
      <td>
        <select class="edit-responsible table-select">
          <option value="">-</option>
          <option value="Олексій Соколов" ${
            issue.responsible === 'Олексій Соколов' ? 'selected' : ''
          }>Олексій Соколов</option>
          <option value="Дмитро Сандул" ${
            issue.responsible === 'Дмитро Сандул' ? 'selected' : ''
          }>Дмитро Сандул</option>
          <option value="Нікіта Чичикало" ${
            issue.responsible === 'Нікіта Чичикало' ? 'selected' : ''
          }>Нікіта Чичикало</option>
          <option value="Ігор Драга" ${issue.responsible === 'Ігор Драга' ? 'selected' : ''}>Ігор Драга</option>
        </select>
      </td>
      <td>
        <select class="edit-importance table-select">
          <option value="Високий" ${issue.importance === 'Високий' ? 'selected' : ''}>Високий</option>
          <option value="Середній" ${issue.importance === 'Середній' ? 'selected' : ''}>Середній</option>
          <option value="Низький" ${issue.importance === 'Низький' ? 'selected' : ''}>Низький</option>
        </select>
      </td>
      <td>
        <button class="btn-small btn-delete" onclick="deleteRow(this)">🗑️</button>
      </td>
    `
    previewBody.appendChild(row)
  })
}

function viewFullDescription(index) {
  const issue = parsedData[index]
  const modal = document.createElement('div')
  modal.className = 'modal'
  modal.innerHTML = `
    <div class="modal-content">
      <div class="modal-header">
        <h3>Повний опис звернення</h3>
        <button class="modal-close" onclick="this.closest('.modal').remove()">✕</button>
      </div>
      <div class="modal-body">
        <pre>${escapeHtml(issue.description)}</pre>
      </div>
      <div class="modal-footer">
        <button class="btn btn-primary" onclick="this.closest('.modal').remove()">Закрити</button>
      </div>
    </div>
  `
  document.body.appendChild(modal)
}

function deleteRow(button) {
  const row = button.closest('tr')
  const index = Array.from(previewBody.children).indexOf(row)

  if (confirm('Видалити це звернення?')) {
    parsedData.splice(index, 1)
    row.remove()
  }
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
