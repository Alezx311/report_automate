const axios = require('axios')

/**
 * GraphApiParser - парсер для роботи з Microsoft Graph API
 */
class GraphApiParser {
  constructor(config) {
    this.config = {
      tenant: config.tenant,
      clientId: config.clientId,
      clientSecret: config.clientSecret,
      user: config.user,
      password: config.password,
    }
    this.token = null
    this.tokenExpiry = null
  }

  /**
   * Отримання токену через ROPC (Resource Owner Password Credentials)
   */
  async getToken() {
    // Перевіряємо чи токен ще дійсний (з запасом 5 хв)
    if (this.token && this.tokenExpiry && Date.now() < this.tokenExpiry - 5 * 60 * 1000) {
      return this.token
    }

    const tokenUrl = `https://login.microsoftonline.com/${this.config.tenant}/oauth2/v2.0/token`

    const params = new URLSearchParams({
      grant_type: 'password',
      client_id: this.config.clientId,
      client_secret: this.config.clientSecret,
      scope: 'https://graph.microsoft.com/.default openid profile email',
      username: this.config.user,
      password: this.config.password,
    })

    try {
      const response = await axios.post(tokenUrl, params, {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        timeout: 10000,
      })

      this.token = response.data.access_token
      // expires_in вказується в секундах
      this.tokenExpiry = Date.now() + response.data.expires_in * 1000

      return this.token
    } catch (error) {
      throw new Error(`Failed to get token: ${error.response?.data?.error_description || error.message}`)
    }
  }

  /**
   * Тестування підключення
   */
  async testConnection() {
    try {
      const token = await this.getToken()

      // Перевіряємо доступ до Graph API
      const response = await axios.get('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: `Bearer ${token}` },
        timeout: 10000,
      })

      return {
        success: true,
        user: response.data.userPrincipalName || response.data.mail,
        displayName: response.data.displayName,
      }
    } catch (error) {
      return {
        success: false,
        error: error.response?.data?.error?.message || error.message,
      }
    }
  }

  /**
   * Отримання списку папок
   */
  async listMailFolders() {
    const token = await this.getToken()
    const folders = []
    let url = 'https://graph.microsoft.com/v1.0/me/mailFolders'

    while (url) {
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        timeout: 30000,
      })

      folders.push(...response.data.value)
      url = response.data['@odata.nextLink']
    }

    return folders
  }

  /**
   * Знаходження ID папки за назвою
   */
  async getFolderIdByName(folderName) {
    const folders = await this.listMailFolders()
    const folder = folders.find(f => f.displayName === folderName)
    return folder ? folder.id : null
  }

  /**
   * Отримання папок зі статистикою
   */
  async getFoldersWithStats() {
    const folders = await this.listMailFolders()

    return folders.map(f => ({
      name: f.displayName,
      id: f.id,
      totalItemCount: f.totalItemCount || 0,
      unreadItemCount: f.unreadItemCount || 0,
      childFolderCount: f.childFolderCount || 0,
    }))
  }

  /**
   * Отримання непрочитаних повідомлень з папки
   */
  async fetchUnreadMessages(folderId, pageSize = 50) {
    const token = await this.getToken()
    const messages = []
    let url = `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}/messages?$filter=isRead ne true&$top=${pageSize}`

    while (url) {
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        timeout: 30000,
      })

      messages.push(...response.data.value)
      url = response.data['@odata.nextLink']
    }

    return messages
  }

  /**
   * Отримання повідомлень з папки за датами
   */
  async fetchEmails(options) {
    const { folder = 'Inbox', startDate, endDate } = options

    // Знаходимо ID папки
    const folderId = await this.getFolderIdByName(folder)
    if (!folderId) {
      throw new Error(`Folder "${folder}" not found`)
    }

    const token = await this.getToken()
    const messages = []

    // Формуємо фільтр
    let filter = ''
    const filters = []

    if (startDate) {
      filters.push(`receivedDateTime ge ${new Date(startDate).toISOString()}`)
    }
    if (endDate) {
      filters.push(`receivedDateTime le ${new Date(endDate).toISOString()}`)
    }

    if (filters.length > 0) {
      filter = `$filter=${filters.join(' and ')}&`
    }

    let url = `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}/messages?${filter}$top=100&$orderby=receivedDateTime desc`

    console.log(`Fetching from folder: ${folder} (${folderId})`)

    while (url) {
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        timeout: 30000,
      })

      const batch = response.data.value
      console.log(`Fetched ${batch.length} messages`)

      // Конвертуємо формат Graph API до нашого внутрішнього формату
      const converted = batch.map(msg => this.convertGraphMessage(msg))
      messages.push(...converted)

      url = response.data['@odata.nextLink']

      // Обмеження для безпеки
      if (messages.length > 10000) {
        console.warn('Reached 10k messages limit')
        break
      }
    }

    return messages
  }

  /**
   * Конвертація повідомлення з формату Graph API до внутрішнього формату
   */
  convertGraphMessage(graphMsg) {
    return {
      messageId: graphMsg.id,
      subject: graphMsg.subject || '(no subject)',
      from: graphMsg.from?.emailAddress?.address || '',
      fromName: graphMsg.from?.emailAddress?.name || '',
      to: graphMsg.toRecipients?.map(r => r.emailAddress?.address).join(', ') || '',
      date: new Date(graphMsg.receivedDateTime),
      body: this.extractBody(graphMsg),
      isRead: graphMsg.isRead,
      hasAttachments: graphMsg.hasAttachments,
      importance: graphMsg.importance,
      conversationId: graphMsg.conversationId,
    }
  }

  /**
   * Вилучення тіла листа
   */
  extractBody(graphMsg) {
    const body = graphMsg.body || {}

    if (body.contentType === 'html') {
      // Видаляємо HTML теги для текстового аналізу
      return this.stripHtml(body.content || '')
    }

    return body.content || ''
  }

  /**
   * Видалення HTML тегів
   */
  stripHtml(html) {
    if (!html) return ''

    // Видаляємо script/style блоки (з прапорцем i для case-insensitive)
    let text = html.replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, '')

    // Видаляємо коментарі
    text = text.replace(/<!--[\s\S]*?-->/g, '')

    // Видаляємо теги
    text = text.replace(/<[^>]+>/g, ' ')

    // Декодуємо HTML entities
    text = text
      .replace(/&nbsp;/g, ' ')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")

    // Видаляємо зайві пробіли
    text = text.replace(/\s+/g, ' ').trim()

    return text
  }

  /**
   * Отримання вкладень повідомлення
   */
  async getMessageAttachments(messageId) {
    const token = await this.getToken()
    const attachments = []
    let url = `https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments`

    while (url) {
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        timeout: 30000,
      })

      attachments.push(...response.data.value)
      url = response.data['@odata.nextLink']
    }

    return attachments
  }

  /**
   * Позначити повідомлення як прочитане
   */
  async markMessageRead(messageId) {
    const token = await this.getToken()

    await axios.patch(
      `https://graph.microsoft.com/v1.0/me/messages/${messageId}`,
      { isRead: true },
      {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        timeout: 10000,
      },
    )
  }

  /**
   * Отримання повного тіла повідомлення в HTML
   */
  async getMessageBodyHtml(messageId) {
    const token = await this.getToken()

    const response = await axios.get(`https://graph.microsoft.com/v1.0/me/messages/${messageId}?$select=body`, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      timeout: 10000,
    })

    const body = response.data.body || {}

    if (body.contentType === 'html') {
      return body.content || ''
    }

    // Якщо plain text, обгортаємо в <pre>
    return `<pre>${body.content || ''}</pre>`
  }

  async connect() {
    // Для сумісності з IMAP інтерфейсом
    await this.getToken()
  }

  async disconnect() {
    // Для сумісності з IMAP інтерфейсом
    // Graph API не потребує явного закриття з'єднання
  }
}

module.exports = GraphApiParser
