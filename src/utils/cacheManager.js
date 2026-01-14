const fs = require('fs').promises
const path = require('path')
const { app } = require('electron')

/**
 * CacheManager - управління кешем завантажених листів
 */
class CacheManager {
  constructor() {
    // Папка для кешу в userData Electron
    this.cacheDir = path.join(app.getPath('userData'), 'mail-cache')
  }

  /**
   * Ініціалізація папки кешу
   */
  async initialize() {
    try {
      await fs.mkdir(this.cacheDir, { recursive: true })
      console.log('Cache directory initialized:', this.cacheDir)
    } catch (error) {
      console.error('Failed to create cache directory:', error)
      throw error
    }
  }

  /**
   * Генерація імені файлу кешу
   */
  generateCacheFileName(source, startDate, endDate, folders = []) {
    const start = startDate ? new Date(startDate).toISOString().split('T')[0] : 'all'
    const end = endDate ? new Date(endDate).toISOString().split('T')[0] : 'now'
    const folderHash = folders.length > 0 ? folders.join('-').substring(0, 30) : 'all'
    const timestamp = Date.now()

    return `${source}_${start}_${end}_${folderHash}_${timestamp}.json`
  }

  /**
   * Збереження даних в кеш
   */
  async saveToCache(source, data, metadata) {
    const fileName = this.generateCacheFileName(
      source,
      metadata.startDate,
      metadata.endDate,
      metadata.folders
    )
    const filePath = path.join(this.cacheDir, fileName)

    const cacheData = {
      source,
      metadata: {
        ...metadata,
        cachedAt: new Date().toISOString(),
        fileName,
      },
      messages: data,
      stats: {
        totalMessages: data.length,
        dateRange: {
          start: metadata.startDate,
          end: metadata.endDate,
        },
        folders: metadata.folders || [],
      },
    }

    await fs.writeFile(filePath, JSON.stringify(cacheData, null, 2), 'utf8')

    console.log(`Cached ${data.length} messages to ${fileName}`)

    return {
      success: true,
      filePath,
      fileName,
      messageCount: data.length,
    }
  }

  /**
   * Отримання списку кешованих файлів
   */
  async listCachedFiles() {
    try {
      const files = await fs.readdir(this.cacheDir)
      const cacheFiles = []

      for (const file of files) {
        if (!file.endsWith('.json')) continue

        const filePath = path.join(this.cacheDir, file)
        const stats = await fs.stat(filePath)

        try {
          // Читаємо метадані з файлу
          const content = await fs.readFile(filePath, 'utf8')
          const data = JSON.parse(content)

          cacheFiles.push({
            fileName: file,
            filePath,
            source: data.source,
            messageCount: data.messages?.length || 0,
            cachedAt: data.metadata?.cachedAt,
            dateRange: data.stats?.dateRange,
            folders: data.stats?.folders || [],
            size: stats.size,
            sizeFormatted: this.formatFileSize(stats.size),
          })
        } catch (error) {
          console.error(`Failed to read cache file ${file}:`, error)
        }
      }

      // Сортуємо за датою створення (новіші спочатку)
      cacheFiles.sort((a, b) => new Date(b.cachedAt) - new Date(a.cachedAt))

      return cacheFiles
    } catch (error) {
      if (error.code === 'ENOENT') {
        await this.initialize()
        return []
      }
      throw error
    }
  }

  /**
   * Завантаження даних з кешу
   */
  async loadFromCache(fileName) {
    const filePath = path.join(this.cacheDir, fileName)

    try {
      const content = await fs.readFile(filePath, 'utf8')
      const data = JSON.parse(content)

      console.log(`Loaded ${data.messages?.length || 0} messages from cache: ${fileName}`)

      return {
        success: true,
        data: data.messages,
        metadata: data.metadata,
        stats: data.stats,
      }
    } catch (error) {
      console.error('Failed to load from cache:', error)
      return {
        success: false,
        error: error.message,
      }
    }
  }

  /**
   * Видалення файлу з кешу
   */
  async deleteCache(fileName) {
    const filePath = path.join(this.cacheDir, fileName)

    try {
      await fs.unlink(filePath)
      console.log(`Deleted cache file: ${fileName}`)
      return { success: true }
    } catch (error) {
      console.error('Failed to delete cache:', error)
      return {
        success: false,
        error: error.message,
      }
    }
  }

  /**
   * Очищення всього кешу
   */
  async clearAllCache() {
    try {
      const files = await fs.readdir(this.cacheDir)

      for (const file of files) {
        if (file.endsWith('.json')) {
          await fs.unlink(path.join(this.cacheDir, file))
        }
      }

      console.log('All cache cleared')
      return { success: true, deletedCount: files.length }
    } catch (error) {
      console.error('Failed to clear cache:', error)
      return {
        success: false,
        error: error.message,
      }
    }
  }

  /**
   * Отримання статистики кешу
   */
  async getCacheStats() {
    try {
      const files = await this.listCachedFiles()

      const totalSize = files.reduce((sum, file) => sum + file.size, 0)
      const totalMessages = files.reduce((sum, file) => sum + file.messageCount, 0)

      return {
        totalFiles: files.length,
        totalSize,
        totalSizeFormatted: this.formatFileSize(totalSize),
        totalMessages,
        files,
      }
    } catch (error) {
      return {
        totalFiles: 0,
        totalSize: 0,
        totalSizeFormatted: '0 B',
        totalMessages: 0,
        files: [],
      }
    }
  }

  /**
   * Форматування розміру файлу
   */
  formatFileSize(bytes) {
    if (bytes === 0) return '0 B'

    const k = 1024
    const sizes = ['B', 'KB', 'MB', 'GB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))

    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i]
  }
}

module.exports = CacheManager
