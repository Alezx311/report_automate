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
   */
  processMessages(messages) {
    console.log(`🔄 Обробка ${messages.length} повідомлень...`)

    // Групування в threads
    const threads = this._groupByThread(messages)
    console.log(`📊 Згруповано в ${Object.keys(threads).length} threads`)

    const issues = []

    Object.values(threads).forEach(thread => {
      // Сортуємо за часом
      thread.sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime))

      let i = 0
      while (i < thread.length) {
        const currentMsg = thread[i]
        const isSupport = this._isFromSupport(currentMsg.senderEmail)

        // Збираємо повідомлення для звернення
        const issueMessages = [currentMsg]
        let j = i + 1
        let foundResponse = false

        while (j < thread.length) {
          const nextMsg = thread[j]
          const nextIsSupport = this._isFromSupport(nextMsg.senderEmail)

          if (nextIsSupport !== isSupport) {
            issueMessages.push(nextMsg)
            foundResponse = true
            break
          }

          issueMessages.push(nextMsg)
          j++
        }

        // Створюємо звернення
        const issue = this._createIssue(issueMessages, isSupport, foundResponse)
        issues.push(issue)

        i = foundResponse ? j + 1 : thread.length
      }
    })

    // Статистика
    const stats = this._calculateStats(issues, Object.keys(threads).length)

    console.log(`✅ Створено ${issues.length} звернень`)

    return { issues, stats }
  }

  /**
   * Обробка задач з Jira
   */
  processJiraIssues(jiraIssues) {
    console.log(`🔄 Обробка ${jiraIssues.length} Jira задач...`)

    const issues = []

    for (const jiraIssue of jiraIssues) {
      // Знаходимо дату переходу в Assigned (початок)
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
        solution: completedTransition ? 'Задача завершена' : '',
        dateResolved: endDate ? this._formatDate(endDate) : '',
        timeResolved: endDate ? this._formatTime(endDate) : '',
        importance: this._mapJiraPriority(jiraIssue.priority),
        source: 'Jira',
      })
    }

    const stats = this._calculateStats(issues, jiraIssues.length)

    console.log(`✅ Оброблено ${issues.length} звернень з Jira`)

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
        { id: 'description', title: 'Опис' },
        { id: 'status', title: 'Статус' },
        { id: 'responsible', title: 'Відповідальний' },
        { id: 'solution', title: 'Рішення' },
        { id: 'dateResolved', title: 'Дата вирішення' },
        { id: 'timeResolved', title: 'Час вирішення' },
        { id: 'importance', title: 'Важливість' },
        { id: 'jiraKey', title: 'Jira Key' },
        { id: 'source', title: 'Джерело' },
      ],
      encoding: 'utf8',
    })

    await csvWriter.writeRecords(issues)

    console.log(`✅ CSV створено: ${filePath}`)

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
    return this.supportEmails.some(supportEmail => email.toLowerCase().includes(supportEmail))
  }

  _createIssue(messages, isFirstFromSupport, hasResponse) {
    const firstMsg = messages[0]
    const lastMsg = messages[messages.length - 1]

    let status = 'У процесі'
    let responsible = ''
    let solution = ''
    let dateResolved = ''
    let timeResolved = ''

    if (hasResponse) {
      const responseMsg = lastMsg

      if (isFirstFromSupport) {
        status = 'Вирішено'
        responsible = this._getResponsibleName(firstMsg.senderEmail)
      } else {
        responsible = this._getResponsibleName(responseMsg.senderEmail)

        const hasConfirmation = messages.some(
          m => m.body.toLowerCase().includes('дякую') || m.body.toLowerCase().includes('працює'),
        )

        status = hasConfirmation ? 'Вирішено' : 'Вирішено частково'
      }

      solution = this._extractSolution(responseMsg.body)
      dateResolved = this._formatDate(responseMsg.receivedDateTime)
      timeResolved = this._formatTime(responseMsg.receivedDateTime)
    } else {
      status = 'Не закрито'
    }

    return {
      dateRegistered: this._formatDate(firstMsg.receivedDateTime),
      timeRegistered: this._formatTime(firstMsg.receivedDateTime),
      system: this._extractSystemFromText(firstMsg.subject + ' ' + firstMsg.body),
      messageCount: messages.length,
      subject: firstMsg.subject,
      description: this._buildDescription(messages),
      status,
      responsible,
      solution,
      dateResolved,
      timeResolved,
      importance: this._calculateImportance(firstMsg.subject + ' ' + firstMsg.body),
      source: 'Email',
    }
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

  _buildDescription(messages) {
    let desc = '=== ПЕРЕПИСКА ===\n\n'

    messages.forEach((msg, i) => {
      const date = this._formatDate(msg.receivedDateTime)
      const time = this._formatTime(msg.receivedDateTime)
      desc += `[${date} ${time}] ${msg.senderName}:\n`
      desc += `${msg.body.substring(0, 500)}\n`

      if (i < messages.length - 1) {
        desc += '\n---\n\n'
      }
    })

    return desc
  }

  _extractSolution(body) {
    return body.substring(0, 500)
  }

  _calculateImportance(text) {
    const lowerText = text.toLowerCase()
    const high = ['прод', 'prod', 'терміново', 'критично', 'urgent', 'critical']
    const low = ['тест', 'test', 'аналіз']

    if (high.some(kw => lowerText.includes(kw))) return 'Високий'
    if (low.some(kw => lowerText.includes(kw))) return 'Низький'
    return 'Середній'
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
      console.error('⚠️ Помилка завантаження систем:', error.message)
      return ['ESB', 'IPS', 'FICO']
    }
  }
}

module.exports = ReportGenerator
