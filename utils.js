// utils.js - Новий файл з допоміжними функціями
const iconv = require('iconv-lite')

function fixEncoding(str) {
  return str
  if (!str) return ''
  const buffer = Buffer.from(str, 'latin1')
  return iconv.decode(buffer, 'win1251')
}

function stripHtml(html) {
  return html
    .replace(/<[^>]*>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

// функція для парсингу дати з body
function parseDateFromBody(body) {
  const datePattern = /Sent:\s*(\w+day),\s*(\w+)\s*(\d+),\s*(\d+)\s*(\d+:\d+\s*[AP]M)/i
  const match = body.match(datePattern)
  if (match) {
    const [, dayOfWeek, month, day, year, time] = match
    const dateStr = `${month} ${day}, ${year} ${time}`
    const parsed = Date.parse(dateStr)
    if (!isNaN(parsed)) {
      return new Date(parsed)
    }
  }
  return null
}

function groupByThread(messages) {
  const threads = {}
  messages.forEach(msg => {
    const convId = msg.conversationId
    if (!threads[convId]) threads[convId] = []
    threads[convId].push(msg)
  })
  return threads
}

function extractSystem(body) {
  const match = body.match(/([A-Z]+_[A-Z]+)/g)
  return match ? match[0] : ''
}

function cleanBody(body) {
  if (!body) return ''
  // Видалити null-байти
  body = body.replace(/\x00/g, '')
  // Видалити заголовки From:, To:, Cc:, Sent: тощо
  body = body.replace(/From:.*?\r?\n/g, '')
  body = body.replace(/To:.*?\r?\n/g, '')
  body = body.replace(/Cc:.*?\r?\n/g, '')
  body = body.replace(/Sent:.*?\r?\n/g, '')
  body = body.replace(/Subject:.*?\r?\n/g, '')
  // Видалити сигнатури (після "З повагою" або "--")
  const signatureIndex = body.search(/(--|З повагою|Best regards)/i)
  if (signatureIndex !== -1) {
    body = body.substring(0, signatureIndex).trim()
  }
  // Видалити зайві рядки та пробіли
  body = body.replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim()
  return body
}

function extractDescription(body, keywords) {
  const cleaned = cleanBody(body)
  const foundKeywords = keywords.filter(keyword => cleaned.toLowerCase().includes(keyword.toLowerCase()))
  let description = cleaned.substring(0, 1000)
  if (foundKeywords.length > 0) {
    description += ' [Ключові слова: ' + foundKeywords.join(', ') + ']'
  }
  return description
}

function extractSolution(body) {
  return cleanBody(body).substring(0, 500)
}

function calculateImportance(body) {
  const highKeywords = ['терміново', 'критично', 'urgent', 'critical', 'високий', 'прод']
  const lowKeywords = ['тест', 'аналіз', 'консультація']
  if (highKeywords.some(kw => body.toLowerCase().includes(kw))) return 'Високий'
  if (lowKeywords.some(kw => body.toLowerCase().includes(kw))) return 'Низький'
  return 'Середній'
}

function formatDate(dateTime) {
  const date = new Date(dateTime)
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

function formatTime(dateTime) {
  const date = new Date(dateTime)
  const hours = String(date.getHours()).padStart(2, '0')
  const minutes = String(date.getMinutes()).padStart(2, '0')
  const seconds = String(date.getSeconds()).padStart(2, '0')
  return `${hours}:${minutes}:${seconds}`
}

// Нова функція: Формування чистого тексту переписки
function formatThreadText(thread) {
  let text = `Всього листів: ${thread.length}\n\n`
  thread.forEach((msg, index) => {
    const sender = msg.senderName || msg.senderEmail || 'Невідомо'
    const time = `${formatDate(msg.receivedDateTime)} ${formatTime(msg.receivedDateTime)}`
    const body = cleanBody(msg.body).substring(0, 200) // Коротко
    text += `Лист ${index + 1}: ${sender} о ${time}: ${body}...\n`
  })
  return text.trim()
}

function processThreads(threads, supportEmailsStr, keywordsStr, filterText) {
  const supportEmails = supportEmailsStr
    .split(',')
    .map(e => e.trim().toLowerCase())
    .filter(e => e)
  const keywords = keywordsStr
    .split(',')
    .map(k => k.trim().toLowerCase())
    .filter(k => k)

  const emailToName = {
    'oleksiy_sokolov@service-team.biz': 'Олексій Соколов',
    'dmytro_sandul@service-team.biz': 'Дмитро Сандул',
    'nikita_chychykalo@service-team.biz': 'Нікіта Чичикало',
    'ihor_draha@service-team.biz': 'Ігор Драга',
  }

  function getResponsible(email) {
    const lowerEmail = email.toLowerCase()
    for (const [key, name] of Object.entries(emailToName)) {
      if (lowerEmail.includes(key)) return name
    }
    return 'Невідомо'
  }

  const issues = []

  Object.values(threads).forEach(thread => {
    if (
      filterText &&
      !thread.some(
        msg =>
          msg.subject.toLowerCase().includes(filterText.toLowerCase()) ||
          msg.body.toLowerCase().includes(filterText.toLowerCase()),
      )
    )
      return

    let openIssue = null

    thread.forEach(msg => {
      const isSupport = supportEmails.some(email => msg.senderEmail.toLowerCase().includes(email))

      if (!isSupport) {
        if (openIssue) {
          openIssue.status = 'Вирішено частково'
          openIssue.responsible = '(клієнт продовжив)'
          openIssue.solution = 'Продовжено переписку'
          openIssue.dateResolved = formatDate(msg.receivedDateTime)
          openIssue.timeResolved = formatTime(msg.receivedDateTime)
          issues.push(openIssue)
        }

        openIssue = {
          dateRegistered: formatDate(msg.receivedDateTime),
          timeRegistered: formatTime(msg.receivedDateTime),
          system: extractSystem(msg.body),
          ticketId: '',
          communication: 'Пошта',
          description: extractDescription(msg.body, keywords),
          status: 'Вирішено',
          responsible: '',
          solution: '',
          dateResolved: '',
          timeResolved: '',
          importance: calculateImportance(msg.body),
          subject: msg.subject,
        }
      } else {
        if (openIssue) {
          openIssue.responsible = getResponsible(msg.senderEmail)
          openIssue.solution = extractSolution(msg.body)
          openIssue.dateResolved = formatDate(msg.receivedDateTime)
          openIssue.timeResolved = formatTime(msg.receivedDateTime)
          const hasThankYou = msg.body.toLowerCase().includes('дякую') || msg.body.toLowerCase().includes('працює')
          openIssue.status = hasThankYou ? 'Вирішено' : 'Вирішено частково'
          issues.push(openIssue)
          openIssue = null
        }
      }
    })

    if (openIssue) {
      openIssue.status = 'Не закрито'
      openIssue.importance = 'Високий'
      issues.push(openIssue)
    }
  })

  const stats = {
    totalThreads: Object.keys(threads).length,
    total: issues.length,
    resolved: issues.filter(i => i.status.startsWith('Вирішено')).length,
    inProgress: issues.filter(i => i.status === 'Не закрито').length,
  }

  return { issues, stats }
}

module.exports = {
  fixEncoding,
  stripHtml,
  groupByThread,
  parseDateFromBody,
  extractSystem,
  cleanBody,
  extractDescription,
  extractSolution,
  calculateImportance,
  formatDate,
  formatTime,
  formatThreadText,
  processThreads,
}
