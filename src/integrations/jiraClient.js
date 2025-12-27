const JiraApi = require('jira-client')

class JiraClient {
  constructor(config) {
    this.jira = new JiraApi({
      protocol: 'https',
      host: config.host.replace('https://', '').replace('http://', ''),
      username: config.email,
      password: config.apiToken,
      apiVersion: '2',
      strictSSL: true,
    })

    this.projectKey = config.projectKey || 'SUPPORT'
    this.config = config
  }

  async testConnection() {
    try {
      const user = await this.jira.getCurrentUser()
      console.log('✅ Jira підключено:', user.displayName)
      return { success: true, user }
    } catch (error) {
      console.error('❌ Помилка підключення до Jira:', error.message)
      return { success: false, error: error.message }
    }
  }

  async getProject() {
    try {
      const project = await this.jira.getProject(this.projectKey)
      return project
    } catch (error) {
      console.error('❌ Помилка отримання проекту:', error.message)
      return null
    }
  }

  /**
   * Завантаження задач з Jira з фільтрацією по статусах
   */
  async fetchIssues(options = {}) {
    try {
      const { startDate, endDate, statuses = ['Assigned', 'In Progress', 'Completed'] } = options

      let jql = `project = ${this.projectKey}`

      if (statuses && statuses.length > 0) {
        const statusFilter = statuses.map(s => `"${s}"`).join(',')
        jql += ` AND status IN (${statusFilter})`
      }

      if (startDate) {
        jql += ` AND created >= "${startDate}"`
      }

      if (endDate) {
        jql += ` AND created <= "${endDate}"`
      }

      jql += ` ORDER BY created DESC`

      console.log('🔍 JQL запит:', jql)

      const searchResults = await this.jira.searchJira(jql, {
        maxResults: 1000,
        fields: [
          'summary',
          'description',
          'status',
          'assignee',
          'created',
          'updated',
          'resolutiondate',
          'priority',
          'issuetype',
        ],
      })

      console.log(`✅ Знайдено ${searchResults.issues.length} задач`)

      // Отримуємо історію переходів для кожної задачі
      const issuesWithHistory = []

      for (const issue of searchResults.issues) {
        try {
          const changelog = await this.jira.getIssueChangelog(issue.key)

          issuesWithHistory.push({
            key: issue.key,
            summary: issue.fields.summary,
            description: issue.fields.description || '',
            status: issue.fields.status.name,
            assignee: issue.fields.assignee?.displayName || 'Не призначено',
            created: issue.fields.created,
            updated: issue.fields.updated,
            resolutionDate: issue.fields.resolutiondate,
            priority: issue.fields.priority?.name || 'Medium',
            issueType: issue.fields.issuetype.name,
            statusHistory: this._extractStatusHistory(changelog),
          })
        } catch (err) {
          console.error(`⚠️ Помилка отримання історії для ${issue.key}:`, err.message)
          issuesWithHistory.push({
            key: issue.key,
            summary: issue.fields.summary,
            description: issue.fields.description || '',
            status: issue.fields.status.name,
            assignee: issue.fields.assignee?.displayName || 'Не призначено',
            created: issue.fields.created,
            statusHistory: [],
          })
        }
      }

      return issuesWithHistory
    } catch (error) {
      console.error('❌ Помилка завантаження задач:', error.message)
      throw error
    }
  }

  /**
   * Створення issue в Jira
   */
  async createIssue(issueData) {
    try {
      const issue = {
        fields: {
          project: { key: this.projectKey },
          summary: issueData.subject || 'Звернення з email',
          description: this._formatDescription(issueData),
          issuetype: { name: issueData.issueType || 'Task' },
          priority: { name: this._mapImportance(issueData.importance) },
          labels: this._generateLabels(issueData),
        },
      }

      if (issueData.responsible) {
        const user = await this.findUserByName(issueData.responsible)
        if (user) {
          issue.fields.assignee = { accountId: user.accountId }
        }
      }

      const result = await this.jira.addNewIssue(issue)
      console.log(`✅ Створено Jira issue: ${result.key}`)

      // Якщо статус "Вирішено", переводимо в Completed
      if (issueData.status === 'Вирішено') {
        await this.transitionIssue(result.key, 'Completed')
      }

      return {
        success: true,
        issueKey: result.key,
        issueUrl: `${this.config.host}/browse/${result.key}`,
      }
    } catch (error) {
      console.error('❌ Помилка створення issue:', error.message)
      return { success: false, error: error.message }
    }
  }

  /**
   * Масове створення issues
   */
  async createBulkIssues(issuesData, onProgress) {
    const results = []

    for (let i = 0; i < issuesData.length; i++) {
      const result = await this.createIssue(issuesData[i])
      results.push(result)

      if (onProgress) {
        onProgress(i + 1, issuesData.length, result)
      }

      await this._delay(500)
    }

    return results
  }

  /**
   * Перехід issue в інший статус
   */
  async transitionIssue(issueKey, toStatus) {
    try {
      const transitions = await this.jira.listTransitions(issueKey)
      const transition = transitions.transitions.find(t => t.to.name.toLowerCase() === toStatus.toLowerCase())

      if (!transition) {
        console.error(`⚠️ Не знайдено перехід до статусу "${toStatus}"`)
        return false
      }

      await this.jira.transitionIssue(issueKey, {
        transition: { id: transition.id },
      })

      console.log(`✅ ${issueKey} -> ${toStatus}`)
      return true
    } catch (error) {
      console.error(`❌ Помилка переходу ${issueKey}:`, error.message)
      return false
    }
  }

  async findUserByName(name) {
    try {
      const users = await this.jira.searchUsers({
        query: name,
        maxResults: 1,
      })
      return users.length > 0 ? users[0] : null
    } catch (error) {
      console.error('⚠️ Помилка пошуку користувача:', error.message)
      return null
    }
  }

  _extractStatusHistory(changelog) {
    const history = []

    if (changelog && changelog.histories) {
      for (const change of changelog.histories) {
        for (const item of change.items) {
          if (item.field === 'status') {
            history.push({
              from: item.fromString,
              to: item.toString,
              date: change.created,
              author: change.author?.displayName || 'Unknown',
            })
          }
        }
      }
    }

    return history
  }

  _formatDescription(issueData) {
    let desc = `h2. Інформація про звернення\n\n`
    desc += `*Дата:* ${issueData.dateRegistered} ${issueData.timeRegistered}\n`
    desc += `*Система:* ${issueData.system}\n`
    desc += `*Листів:* ${issueData.messageCount || 1}\n`
    desc += `*Важливість:* ${issueData.importance}\n\n`
    desc += `h2. Опис\n\n${issueData.description}\n`

    if (issueData.solution) {
      desc += `\nh2. Рішення\n\n${issueData.solution}\n`
    }

    return desc
  }

  _mapImportance(importance) {
    const map = {
      Високий: 'High',
      Середній: 'Medium',
      Низький: 'Low',
    }
    return map[importance] || 'Medium'
  }

  _generateLabels(issueData) {
    const labels = ['email-import']
    if (issueData.system) labels.push(issueData.system.toLowerCase())
    return labels
  }

  _delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms))
  }
}

module.exports = JiraClient
