const fs = require('fs')
const path = require('path')
const createCsvWriter = require('csv-writer').createObjectCsvWriter

class ReportGenerator {
  constructor(options = {}) {
    this.supportEmails = (options.supportEmails || '')
      .split(',')
      .map(e => e.trim().toLowerCase())
      .filter(e => e)

    this.keywords = (options.keywords || '')
      .split(',')
      .map(k => k.trim().toLowerCase())
      .filter(k => k)

    this.systemsList = this._loadSystems()

    this.emailToName = {
      'oleksiy_sokolov@service-team.biz': 'Олексій Соколов',
      'dmytro_sandul@service-team.biz': 'Дмитро Сандул',
      'nikita_chychykalo@service-team.biz': 'Нікіта Чичикало',
      'ihor_draha@service-team.biz': 'Ігор Драга',
    }
  }

  /**
   * Обробка повідомлень з Email (PST або IMAP)
   *
   * Логіка: Кожна відповідь від техпідтримки в thread = окремий запис у звіті
   */
  processMessages(messages) {
    console.log(`Processing ${messages.length} messages...`)

    // Group into threads
    const threads = this._groupByThread(messages)
    console.log(`Grouped into ${Object.keys(threads).length} threads`)

    // DEBUG: показати supportEmails
    console.log('DEBUG: supportEmails для перевірки:', this.supportEmails)

    const issues = []

    Object.values(threads).forEach((thread, threadIdx) => {
      // Sort by time
      thread.sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime))

      // DEBUG: показати перший thread
      if (threadIdx === 0) {
        console.log('DEBUG: First thread messages:')
        thread.forEach((msg, idx) => {
          console.log(`  ${idx}. senderEmail="${msg.senderEmail}", isSupport=${this._isFromSupport(msg.senderEmail)}`)
        })
      }

      // Check if thread has any support responses
      const supportResponses = thread.filter(msg => this._isFromSupport(msg.senderEmail))
      if (supportResponses.length === 0) {
        // Skip threads without support responses
        return
      }

      // Create ONE issue per support response (correct logic)
      supportResponses.forEach(responseMsg => {
        // Find the request message (last message before this support response)
        const responseIdx = thread.indexOf(responseMsg)
        const requestMsg = responseIdx > 0 ? thread[responseIdx - 1] : thread[0]

        const issue = this._createIssueFromResponse(requestMsg, responseMsg, thread)
        issues.push(issue)
      })
    })

    // Statistics
    const stats = this._calculateStats(issues, Object.keys(threads).length)

    console.log(`Created ${issues.length} issues`)

    return { issues, stats }
  }

  /**
   * Обробка задач з Jira
   */
  processJiraIssues(jiraIssues) {
    console.log(`Processing ${jiraIssues.length} Jira issues...`)

    const issues = []

    for (const jiraIssue of jiraIssues) {
      // Find Assigned transition date (start)
      const assignedTransition = jiraIssue.statusHistory.find(h => h.to === 'Assigned')
      const completedTransition = jiraIssue.statusHistory.find(h => h.to === 'Completed')

      const startDate = assignedTransition ? new Date(assignedTransition.date) : new Date(jiraIssue.created)
      const endDate = completedTransition ? new Date(completedTransition.date) : null

      issues.push({
        jiraKey: jiraIssue.key,
        dateRegistered: this._formatDate(startDate),
        timeRegistered: this._formatTime(startDate),
        system: this._extractSystemFromText(jiraIssue.summary + ' ' + jiraIssue.description),
        messageCount: jiraIssue.statusHistory.length + 1,
        subject: jiraIssue.summary,
        description: jiraIssue.description,
        status: this._mapJiraStatus(jiraIssue.status),
        responsible: jiraIssue.assignee,
        solution: completedTransition ? 'Task completed' : '',
        dateResolved: endDate ? this._formatDate(endDate) : '',
        timeResolved: endDate ? this._formatTime(endDate) : '',
        importance: this._mapJiraPriority(jiraIssue.priority),
        source: 'Jira',
      })
    }

    const stats = this._calculateStats(issues, jiraIssues.length)

    console.log(`Processed ${issues.length} issues from Jira`)

    return { issues, stats }
  }

  /**
   * Генерація CSV файлу
   */
  async generateCSV(issues) {
    const outputDir = path.join(__dirname, '../../output')

    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true })
    }

    const timestamp = new Date().toISOString().replace(/:/g, '-').split('.')[0]
    const fileName = `report_${timestamp}.csv`
    const filePath = path.join(outputDir, fileName)

    const csvWriter = createCsvWriter({
      path: filePath,
      header: [
        { id: 'dateRegistered', title: 'Дата реєстрації' },
        { id: 'timeRegistered', title: 'Час реєстрації' },
        { id: 'system', title: 'Система' },
        { id: 'messageCount', title: 'Кількість листів' },
        { id: 'subject', title: 'Тема' },
        { id: 'problemType', title: 'Тип проблеми' },
        { id: 'description', title: 'Опис' },
        { id: 'requestText', title: 'Текст запиту' },
        { id: 'responseText', title: 'Текст відповіді' },
        { id: 'threadPosition', title: 'Позиція в thread' },
        { id: 'status', title: 'Статус' },
        { id: 'responsible', title: 'Відповідальний' },
        { id: 'solution', title: 'Рішення' },
        { id: 'dateResolved', title: 'Дата вирішення' },
        { id: 'timeResolved', title: 'Час вирішення' },
        { id: 'importance', title: 'Важливість' },
        { id: 'conversationHistory', title: 'Історія переписки' },
        { id: 'jiraKey', title: 'Jira Key' },
        { id: 'source', title: 'Джерело' },
      ],
      encoding: 'utf8',
    })

    await csvWriter.writeRecords(issues)

    console.log(`CSV created: ${filePath}`)

    return filePath
  }

  // ============================================
  // Допоміжні методи
  // ============================================

  _groupByThread(messages) {
    const threads = {}

    messages.forEach(msg => {
      const id = msg.conversationId || msg.subject || 'unknown'
      if (!threads[id]) threads[id] = []
      threads[id].push(msg)
    })

    return threads
  }

  _isFromSupport(email) {
    if (!email) return false
    const lowerEmail = email.toLowerCase()
    return this.supportEmails.some(supportEmail => lowerEmail.includes(supportEmail.toLowerCase()))
  }

  /**
   * Створення запису на основі відповіді техпідтримки
   *
   * @param {Object} requestMsg - Повідомлення з запитом (від клієнта)
   * @param {Object} responseMsg - Відповідь техпідтримки
   * @param {Array} thread - Весь thread для контексту
   */
  _createIssueFromResponse(requestMsg, responseMsg, thread) {
    const system = this._extractSystemFromText(requestMsg.subject + ' ' + requestMsg.body)
    const responsible = this._getResponsibleName(responseMsg.senderEmail)

    // Find position of this response in thread
    const supportResponses = thread.filter(msg => this._isFromSupport(msg.senderEmail))
    const responseIdx = supportResponses.indexOf(responseMsg)
    const threadPosition = `${responseIdx + 1} з ${supportResponses.length}`

    // Build conversation history
    const conversationHistory = this._buildConversationHistory(thread)

    // Detect problem type from request
    const problemType = this._detectProblemType(requestMsg.body)

    // Extract request and response texts
    const requestText = this._extractShortDescription(requestMsg.body)
    const responseText = this._extractSolution(responseMsg.body)

    return {
      dateRegistered: this._formatDate(requestMsg.receivedDateTime),
      timeRegistered: this._formatTime(requestMsg.receivedDateTime),
      system,
      messageCount: thread.length,
      subject: requestMsg.subject,
      description: this._extractShortDescription(requestMsg.body),
      problemType,
      requestText,
      responseText,
      conversationHistory,
      threadPosition,
      status: 'Вирішено',
      responsible,
      solution: this._extractSolution(responseMsg.body),
      dateResolved: this._formatDate(responseMsg.receivedDateTime),
      timeResolved: this._formatTime(responseMsg.receivedDateTime),
      importance: this._calculateImportance(requestMsg.subject + ' ' + requestMsg.body),
      source: 'Email',
    }
  }

  _extractShortDescription(body) {
    if (!body) return ''

    // Спочатку очищаємо від email garbage (попередження, підписи, headers)
    let cleaned = this._cleanEmailBody(body)

    // Видаляємо HTML теги та CSS стилі
    cleaned = cleaned
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '') // Видаляємо <style> блоки
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '') // Видаляємо <script> блоки
      .replace(/<[^>]+>/g, ' ') // Видаляємо всі HTML теги
      .replace(/&nbsp;/g, ' ') // Замінюємо &nbsp; на пробіл
      .replace(/&[a-z]+;/gi, ' ') // Видаляємо всі HTML entities
      .replace(/&#\d+;/g, ' ') // Видаляємо числові entities
      .replace(/\r?\n/g, ' ') // Замінюємо переноси на пробіли
      .trim()

    // Видаляємо VML namespace символи (v\:, o\:, w\:) - ПЕРЕД видаленням CSS
    cleaned = cleaned
      .replace(/[a-z]\\:/gi, '') // будь-які літери з \: (наприклад v\:, o\:, w\:, m\:, p\:)
      .replace(/[a-z]:/gi, (match) => {
        // Залишаємо тільки двокрапки після українських слів, видаляємо XML namespaces
        return /[а-яА-ЯіІїЇєЄґҐ]:/.test(match) ? match : ''
      })

    // Видаляємо CSS властивості та класи
    cleaned = cleaned
      .replace(/[a-z\-]+:\s*[^;]+;/gi, ' ') // CSS властивості
      .replace(/\.[a-z][a-z0-9\-_]*/gi, ' ') // CSS класи (.className)
      .replace(/#[a-z][a-z0-9\-_]*/gi, ' ') // CSS ID (#elementId)
      .replace(/[{}]/g, ' ') // Фігурні дужки

    // Видаляємо RTF команди та спецсимволи
    cleaned = cleaned
      .replace(/\\[a-z]+\d*/gi, ' ') // RTF команди (\par, \b0, etc)
      .replace(/[vow]\\?\*+/gi, '') // v\*, o\*, w\*, v*, o*, w* символи
      .replace(/\\?\*+[vow]?/gi, '') // \*v, \*o, \*w, *v, *o, *w символи
      .replace(/\.shape/gi, ' ') // .shape залишки
      .replace(/\s[a-z]\\?\*+\s/gi, ' ') // будь-які літери з зірочками

    // Видаляємо не-ASCII та спецсимволи (залишаємо тільки кирилицю, латиницю, цифри та базову пунктуацію)
    cleaned = cleaned.replace(/[^\u0020-\u007E\u0400-\u04FF\u0100-\u017F]/g, ' ')

    // Видаляємо контрольні символи
    cleaned = cleaned.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]/g, '')

    // Очищаємо множинні пробіли
    cleaned = cleaned.replace(/\s+/g, ' ').trim()

    // Якщо після очищення текст порожній або дуже короткий, спробуємо взяти перші слова
    if (cleaned.length < 10 && body.length > 20) {
      // Можливо текст був повністю в спецформаті, спробуємо інший підхід
      const words = body.split(/\s+/).filter(word => {
        // Залишаємо тільки слова що містять букви
        return /[а-яА-ЯіІїЇєЄґҐa-zA-Z]{3,}/.test(word)
      })
      cleaned = words.slice(0, 20).join(' ')
    }

    // Обрізаємо до потрібної довжини
    const maxLength = 200
    if (cleaned.length <= maxLength) return cleaned
    return cleaned.substring(0, maxLength) + '...'
  }

  _getResponsibleName(email) {
    const lowerEmail = email.toLowerCase()
    for (const [key, name] of Object.entries(this.emailToName)) {
      if (lowerEmail.includes(key)) return name
    }
    return 'Невідомо'
  }

  _extractSystemFromText(text) {
    const upperText = text.toUpperCase()

    for (const system of this.systemsList) {
      if (upperText.includes(system.toUpperCase())) {
        return system
      }
    }

    return 'ESB'
  }

  _extractSolution(body) {
    if (!body) return ''

    // Спочатку очищаємо від email garbage (попередження, підписи, headers)
    let cleaned = this._cleanEmailBody(body)

    // Видаляємо HTML теги та CSS стилі
    cleaned = cleaned
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '') // Видаляємо <style> блоки
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '') // Видаляємо <script> блоки
      .replace(/<[^>]+>/g, ' ') // Видаляємо всі HTML теги
      .replace(/&nbsp;/g, ' ') // Замінюємо &nbsp; на пробіл
      .replace(/&[a-z]+;/gi, ' ') // Видаляємо всі HTML entities
      .replace(/&#\d+;/g, ' ') // Видаляємо числові entities
      .replace(/\r?\n/g, ' ') // Замінюємо переноси на пробіли
      .trim()

    // Видаляємо VML namespace символи (v\:, o\:, w\:) - ПЕРЕД видаленням CSS
    cleaned = cleaned
      .replace(/[a-z]\\:/gi, '') // будь-які літери з \: (наприклад v\:, o\:, w\:, m\:, p\:)
      .replace(/[a-z]:/gi, (match) => {
        // Залишаємо тільки двокрапки після українських слів, видаляємо XML namespaces
        return /[а-яА-ЯіІїЇєЄґҐ]:/.test(match) ? match : ''
      })

    // Видаляємо CSS властивості та класи
    cleaned = cleaned
      .replace(/[a-z\-]+:\s*[^;]+;/gi, ' ') // CSS властивості
      .replace(/\.[a-z][a-z0-9\-_]*/gi, ' ') // CSS класи (.className)
      .replace(/#[a-z][a-z0-9\-_]*/gi, ' ') // CSS ID (#elementId)
      .replace(/[{}]/g, ' ') // Фігурні дужки

    // Видаляємо RTF команди та спецсимволи
    cleaned = cleaned
      .replace(/\\[a-z]+\d*/gi, ' ') // RTF команди (\par, \b0, etc)
      .replace(/[vow]\\?\*+/gi, '') // v\*, o\*, w\*, v*, o*, w* символи
      .replace(/\\?\*+[vow]?/gi, '') // \*v, \*o, \*w, *v, *o, *w символи
      .replace(/\.shape/gi, ' ') // .shape залишки
      .replace(/\s[a-z]\\?\*+\s/gi, ' ') // будь-які літери з зірочками

    // Видаляємо не-ASCII та спецсимволи (залишаємо тільки кирилицю, латиницю, цифри та базову пунктуацію)
    cleaned = cleaned.replace(/[^\u0020-\u007E\u0400-\u04FF\u0100-\u017F]/g, ' ')

    // Видаляємо контрольні символи
    cleaned = cleaned.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]/g, '')

    // Очищаємо множинні пробіли
    cleaned = cleaned.replace(/\s+/g, ' ').trim()

    // Якщо після очищення текст порожній або дуже короткий, спробуємо взяти перші слова
    if (cleaned.length < 10 && body.length > 20) {
      const words = body.split(/\s+/).filter(word => {
        return /[а-яА-ЯіІїЇєЄґҐa-zA-Z]{3,}/.test(word)
      })
      cleaned = words.slice(0, 30).join(' ')
    }

    // Обрізаємо до потрібної довжини
    const maxLength = 300
    if (cleaned.length <= maxLength) return cleaned
    return cleaned.substring(0, maxLength) + '...'
  }

  _calculateImportance(text) {
    const lowerText = text.toLowerCase()
    const high = ['прод', 'prod', 'терміново', 'критично', 'urgent', 'critical']
    const low = ['тест', 'test', 'аналіз']

    if (high.some(kw => lowerText.includes(kw))) return 'Високий'
    if (low.some(kw => lowerText.includes(kw))) return 'Низький'
    return 'Середній'
  }

  /**
   * Очищення тексту листа від зайвої інформації
   */
  _cleanEmailBody(body) {
    if (!body) return ''

    let cleaned = body

    // 0. СПОЧАТКУ видаляємо HTML та CSS (найважливіше!)
    cleaned = cleaned
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '') // <style> блоки
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '') // <script> блоки
      .replace(/<[^>]+>/g, ' ') // Всі HTML теги
      .replace(/&nbsp;/g, ' ') // &nbsp;
      .replace(/&[a-z]+;/gi, ' ') // HTML entities
      .replace(/&#\d+;/g, ' ') // Числові entities
      .replace(/&#x[0-9a-f]+;/gi, ' ') // Hex entities

    // 0.1. Видаляємо VML CSS декларації
    cleaned = cleaned
      .replace(/[vow]\\:\*\s*\{[^}]*\}/gi, '') // v\:* {...}, o\:* {...}, w\:* {...}
      .replace(/\.shape\s*\{[^}]*\}/gi, '') // .shape {...}
      .replace(/\{behavior:url\([^)]*\);\}/gi, '') // {behavior:url(...);}
      .replace(/behavior:\s*url\([^)]*\)/gi, '') // behavior:url(...)

    // 0.2. Видаляємо VML namespace символи
    cleaned = cleaned
      .replace(/[vow]\\:\*/gi, '') // v\:*, o\:*, w\:*
      .replace(/[a-z]\\:/gi, '') // будь-які літери з \:

    // 0.3. Видаляємо пошкоджені Unicode символи
    cleaned = cleaned.replace(/[\ud800-\udfff]/g, '') // Surrogate pairs (поганий Unicode)

    // 0.4. Видаляємо null-байти та інший binary garbage
    cleaned = cleaned.replace(/\u0000/g, '') // Null bytes
    cleaned = cleaned.replace(/[\u0001-\u0008\u000B\u000C\u000E-\u001F]/g, '') // Control characters

    // 0.5. Видаляємо послідовності незрозумілих non-ASCII символів
    // Якщо в слові більше 3 підряд символів не латиниці/кирилиці/цифр - видаляємо все слово
    cleaned = cleaned.replace(/\b[^\s]*[^\x00-\x7F\u0400-\u04FF\u0020-\u007E]{4,}[^\s]*\b/g, '')

    // 0.6. Видаляємо залишки RTF та спеціальних символів
    cleaned = cleaned.replace(/\\u\d{4}/g, '') // \uXXXX sequences
    cleaned = cleaned.replace(/\\x[0-9a-f]{2}/gi, '') // \xXX sequences

    // 0.7. Видаляємо HTML entity залишки (nbsp, quot, amp, lt, gt тощо)
    cleaned = cleaned.replace(/\b(nbsp|quot|amp|lt|gt|apos)\b/gi, ' ')
    cleaned = cleaned.replace(/[;:]p>/gi, '') // Залишки типу ":p>" або ";p>"
    cleaned = cleaned.replace(/o:p>/gi, '') // Залишки o:p>

    // 1. Видаляємо попередження про зовнішній лист
    cleaned = cleaned.replace(
      /УВАГА!\s*[–-]\s*Зовнішній лист:.*?не очікували цього листа\.\s*/gi,
      ''
    )

    // 2. Видаляємо email headers (From:, Sent:, To:, Cc:, Subject:)
    cleaned = cleaned.replace(/From:\s*[^\n]+\n/gi, '')
    cleaned = cleaned.replace(/Sent:\s*[^\n]+\n/gi, '')
    cleaned = cleaned.replace(/To:\s*[^\n]+\n/gi, '')
    cleaned = cleaned.replace(/Cc:\s*[^\n]+\n/gi, '')
    cleaned = cleaned.replace(/Subject:\s*[^\n]+\n/gi, '')

    // 3. Видаляємо email підписи (телефони, email, посади)
    cleaned = cleaned.replace(/Phone\s*:\s*\+?[\d\s()]+/gi, '')
    cleaned = cleaned.replace(/Email\s*:\s*[\w.+-]+@[\w.-]+/gi, '')
    cleaned = cleaned.replace(/Website\s*:\s*[\w.:/-]+/gi, '')
    cleaned = cleaned.replace(/тел\.?\s*:\s*\+?[\d\s()]+/gi, '')
    cleaned = cleaned.replace(/тел\.?\s*внутрішній\s*:\s*\d+/gi, '')
    cleaned = cleaned.replace(/моб\.?\s*тел\.?\s*:\s*\+?[\d\s()]+/gi, '')

    // 4. Видаляємо імена з підписів (Dmytro Sandul, Nikita Chychykalo тощо)
    cleaned = cleaned.replace(/\b(Dmytro|Nikita|Dmytro_Sandul|Молойко|Руденко|Міщевський|Антушевич|Дмитренко|Мохамед|Лур'є)\b[^\n]*/gi, '')

    // 5. Видаляємо типові посади
    cleaned = cleaned.replace(/Technical\s+Support\s+Manager/gi, '')
    cleaned = cleaned.replace(/Головний\s+фахівець[^\n]*/gi, '')
    cleaned = cleaned.replace(/\b(Manager|Developer|Head|Керівник|Менеджер|фахівець)\b/gi, '')

    // 6. Видаляємо посади та контактну інформацію з підписів
    cleaned = cleaned.replace(
      /[A-Z][a-z]+\s+[A-Z][a-z]+\s*\|\s*[\w\s]+(?:Manager|Developer|Head|Керівник|Менеджер)[^\n]*/gi,
      ''
    )

    // 7. Видаляємо лінії-розділювачі
    cleaned = cleaned.replace(/__{10,}/g, '')
    cleaned = cleaned.replace(/_{5,}/g, '')

    // 8. Видаляємо стандартні фрази "З повагою"
    cleaned = cleaned.replace(/З повагою,?\s*[\w\s]+/gi, '')

    // 9. Видаляємо адреси
    cleaned = cleaned.replace(/вул\.\s*[\w\s,]+\d+/gi, '')
    cleaned = cleaned.replace(/Київ,\s*\d{5}/gi, '')
    cleaned = cleaned.replace(/Україна/gi, '')

    // 10. Видаляємо цитування попередніх листів (після "From:" або ">")
    const lines = cleaned.split('\n')
    const filteredLines = []
    let inQuote = false

    for (const line of lines) {
      // Якщо знайшли "From:" - це початок цитування
      if (/^From:\s*/i.test(line)) {
        inQuote = true
        continue
      }
      // Якщо рядок починається з ">" - це цитування
      if (/^\s*>/.test(line)) {
        continue
      }
      // Якщо не в цитуванні - додаємо рядок
      if (!inQuote) {
        filteredLines.push(line)
      }
    }

    cleaned = filteredLines.join('\n')

    // 11. Очищаємо множинні порожні рядки
    cleaned = cleaned.replace(/\n{3,}/g, '\n\n')

    // 12. Видаляємо множинні пробіли
    cleaned = cleaned.replace(/\s{3,}/g, ' ')

    // 13. Обрізаємо зайві пробіли
    cleaned = cleaned.trim()

    return cleaned
  }

  /**
   * Побудова повної історії переписки для thread
   */
  _buildConversationHistory(thread) {
    const history = thread.map((msg, idx) => {
      const date = this._formatDate(msg.receivedDateTime)
      const time = this._formatTime(msg.receivedDateTime)
      const from = msg.senderName || msg.senderEmail || 'Unknown'

      // Очищаємо тіло листа від зайвої інформації
      let cleanedBody = this._cleanEmailBody(msg.body)

      // Обрізаємо до 300 символів для історії
      if (cleanedBody.length > 300) {
        cleanedBody = cleanedBody.substring(0, 300) + '...'
      }

      // Якщо після очищення нічого не залишилось - пропускаємо
      if (cleanedBody.length < 10) {
        return null
      }

      return `[${idx + 1}] ${date} ${time} - ${from}:\n${cleanedBody}`
    })

    // Фільтруємо null значення
    const filteredHistory = history.filter(item => item !== null)

    return filteredHistory.join('\n\n---\n\n')
  }

  /**
   * Визначення типу проблеми на основі тексту запиту
   */
  _detectProblemType(text) {
    const lowerText = text.toLowerCase()

    // Шаблони для розпізнавання типів проблем
    const patterns = [
      { keywords: ['лог', 'логи', 'логів', 'log'], type: 'Запит логів' },
      { keywords: ['перезавантажити', 'перезапустити', 'reboot', 'restart'], type: 'Перезавантаження' },
      { keywords: ['помилка', 'error', 'exception', 'failed'], type: 'Помилка' },
      { keywords: ['не працює', 'не відправля', 'не передає', 'not working'], type: 'Збій роботи' },
      { keywords: ['квитанці', 'квитанція', 'receipt'], type: 'Квитанції' },
      { keywords: ['аналіз', 'перевірка', 'analysis', 'check'], type: 'Аналіз/перевірка' },
      { keywords: ['налаштування', 'конфігурація', 'config', 'setup'], type: 'Налаштування' },
      { keywords: ['доступ', 'права', 'access', 'permission'], type: 'Доступ/права' },
      { keywords: ['інтеграція', 'integration'], type: 'Інтеграція' },
      { keywords: ['ранкова перевірка', 'morning check'], type: 'Моніторинг' },
    ]

    for (const pattern of patterns) {
      if (pattern.keywords.some(kw => lowerText.includes(kw))) {
        return pattern.type
      }
    }

    return 'Інше'
  }

  _mapJiraStatus(status) {
    const map = {
      Assigned: 'У процесі',
      'In Progress': 'У процесі',
      Completed: 'Вирішено',
    }
    return map[status] || status
  }

  _mapJiraPriority(priority) {
    const map = {
      High: 'Високий',
      Highest: 'Високий',
      Medium: 'Середній',
      Low: 'Низький',
      Lowest: 'Низький',
    }
    return map[priority] || 'Середній'
  }

  _calculateStats(issues, totalThreads) {
    return {
      totalThreads,
      total: issues.length,
      resolved: issues.filter(i => i.status === 'Вирішено').length,
      partial: issues.filter(i => i.status === 'Вирішено частково').length,
      inProgress: issues.filter(i => i.status === 'У процесі' || i.status === 'Не закрито').length,
      avgMessagesPerIssue:
        issues.length > 0
          ? (issues.reduce((sum, i) => sum + (i.messageCount || 0), 0) / issues.length).toFixed(1)
          : '0',
    }
  }

  _formatDate(date) {
    const d = new Date(date)
    const year = d.getFullYear()
    const month = String(d.getMonth() + 1).padStart(2, '0')
    const day = String(d.getDate()).padStart(2, '0')
    return `${year}-${month}-${day}`
  }

  _formatTime(date) {
    const d = new Date(date)
    const hours = String(d.getHours()).padStart(2, '0')
    const minutes = String(d.getMinutes()).padStart(2, '0')
    const seconds = String(d.getSeconds()).padStart(2, '0')
    return `${hours}:${minutes}:${seconds}`
  }

  _loadSystems() {
    try {
      const systemsPath = path.join(__dirname, '../../config/systems.json')
      return JSON.parse(fs.readFileSync(systemsPath, 'utf8'))
    } catch (error) {
      console.error('Systems loading error:', error.message)
      return ['ESB', 'IPS', 'FICO']
    }
  }
}

module.exports = ReportGenerator
