/**
 * Утиліта для очищення тексту від HTML, VML розмітки та garbage символів
 */

class TextCleaner {
  /**
   * Базове очищення - видаляє HTML/VML теги, headers, підписи
   * Працює коректно для нормальних PST файлів з правильним кодуванням
   */
  static basicClean(text) {
    if (!text) return ''

    let cleaned = text

    // 0. Попередня обробка HTML entities та спеціальних символів
    cleaned = cleaned
      .replace(/&nbsp;/gi, ' ') // Non-breaking space
      .replace(/&lt;/g, '<') // Less than
      .replace(/&gt;/g, '>') // Greater than
      .replace(/&amp;/g, '&') // Ampersand
      .replace(/&quot;/g, '"') // Quote
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

    // 0.3. Видаляємо пошкоджені Unicode символи (тільки surrogate pairs)
    cleaned = cleaned.replace(/[\ud800-\udfff]/g, '')

    // 0.4. Видаляємо null-байти та control characters
    cleaned = cleaned.replace(/\u0000/g, '')
    cleaned = cleaned.replace(/[\u0001-\u0008\u000B\u000C\u000E-\u001F]/g, '')

    // 1. Видаляємо попередження про зовнішній лист
    cleaned = cleaned.replace(/УВАГА!\s*[–-]\s*Зовнішній лист:.*?не очікували цього листа\.\s*/gi, '')

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

    // 4. Видаляємо імена з підписів
    cleaned = cleaned.replace(
      /\b(Dmytro|Nikita|Dmytro_Sandul|Молойко|Руденко|Міщевський|Антушевич|Дмитренко|Мохамед|Лур'є)\b[^\n]*/gi,
      '',
    )

    // 5. Видаляємо типові посади
    cleaned = cleaned.replace(/Technical\s+Support\s+Manager/gi, '')
    cleaned = cleaned.replace(/Головний\s+фахівець[^\n]*/gi, '')
    cleaned = cleaned.replace(/\b(Manager|Developer|Head|Керівник|Менеджер|фахівець)\b/gi, '')

    // 6. Видаляємо посади та контактну інформацію з підписів
    cleaned = cleaned.replace(
      /[A-Z][a-z]+\s+[A-Z][a-z]+\s*\|\s*[\w\s]+(?:Manager|Developer|Head|Керівник|Менеджер)[^\n]*/gi,
      '',
    )

    // 7. Видаляємо лінії-розділювачі
    cleaned = cleaned.replace(/__{10,}/g, '')
    cleaned = cleaned.replace(/_{5,}/g, '')
    cleaned = cleaned.replace(/={5,}/g, '')
    cleaned = cleaned.replace(/-{5,}/g, '')

    // 8. Видаляємо зайві пробіли та порожні рядки
    cleaned = cleaned.replace(/[ \t]+/g, ' ') // Multiple spaces to single
    cleaned = cleaned.replace(/\n{3,}/g, '\n\n') // Max 2 newlines
    cleaned = cleaned.trim()

    return cleaned
  }

  /**
   * Агресивне очищення - для пошкоджених PST файлів з неправильним кодуванням
   * Видаляє всі garbage символи, змішані латиниця+кирилиця слова
   */
  static aggressiveClean(text) {
    if (!text) return ''

    // Спочатку базове очищення
    let cleaned = this.basicClean(text)

    // 0.5. Видаляємо китайські/японські/корейські символи (CJK)
    cleaned = cleaned.replace(/[\u4E00-\u9FFF\u3040-\u309F\u30A0-\u30FF\uAC00-\uD7AF]/g, '')

    // 0.6. Видаляємо всі інші non-printable та специфічні Unicode діапазони
    cleaned = cleaned.replace(
      /[\u0080-\u024F\u0370-\u03FF\u0590-\u05FF\u0600-\u06FF\u0900-\u097F\u0A00-\u0DFF\u1000-\u109F\u1100-\u11FF\u1200-\u137F\u1400-\u167F\u1680-\u169F]/g,
      '',
    )

    // 0.7. Видаляємо послідовності незрозумілих non-ASCII символів
    cleaned = cleaned.replace(/\b[^\s]*[^\x00-\x7F\u0400-\u04FF\u0020-\u007E]{2,}[^\s]*\b/g, '')

    // 0.7.1. Агресивно видаляємо БУДЬ-ЯКІ символи, які не є:
    // - Основною латиницею (A-Z, a-z)
    // - Українською кирилицею (А-Я, а-я, Ґ, ґ, Є, є, І, і, Ї, ї)
    // - Цифрами (0-9)
    // - Базовими розділовими знаками
    cleaned = cleaned.replace(/[^A-Za-z0-9А-Яа-яҐґЄєІіЇї \n\r\t.,;:!?()\-_"'@/+=&#%*<>]/g, '')

    // 0.7.2. Видаляємо слова з змішаною латиницею + кирилицею (garbage)
    // Pattern 1: латиниця → кирилиця (приклад: "sй", "mм", "h-м2")
    cleaned = cleaned.replace(/[A-Za-z]+[\-_]?[А-Яа-яҐґЄєІіЇї]+[A-Za-zА-Яа-яҐґЄєІіЇї0-9\-_]*/g, ' ')
    // Pattern 2: кирилиця → латиниця (приклад: "оs", "ам", "ntgr_status")
    cleaned = cleaned.replace(/[А-Яа-яҐґЄєІіЇї]+[\-_]?[A-Za-z]+[A-Za-zА-Яа-яҐґЄєІіЇї0-9\-_]*/g, ' ')

    // 0.8. Видаляємо залишки RTF та спеціальних символів
    cleaned = cleaned.replace(/\\u\d{4}/g, '')
    cleaned = cleaned.replace(/\\x[0-9a-f]{2}/gi, '')

    // 0.9. Видаляємо HTML entity залишки
    cleaned = cleaned.replace(/\b(nbsp|quot|amp|lt|gt|apos)\b/gi, ' ')
    cleaned = cleaned.replace(/[;:]p>/gi, '')
    cleaned = cleaned.replace(/o:p>/gi, '')

    // Фінальне очищення пробілів
    cleaned = cleaned.replace(/[ \t]+/g, ' ')
    cleaned = cleaned.replace(/\n{3,}/g, '\n\n')
    cleaned = cleaned.trim()

    return cleaned
  }

  /**
   * Головна функція очищення - вибирає режим в залежності від прапорця
   */
  static clean(text, useAggressiveMode = false) {
    if (useAggressiveMode) {
      return this.aggressiveClean(text)
    } else {
      return this.basicClean(text)
    }
  }
}

module.exports = TextCleaner
