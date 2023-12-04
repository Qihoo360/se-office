import Path from 'path-browserify'
import { Plugin } from './plugin'

class CStore {
  getRecentFiles(type: string): any[] {
    const jsonStr = localStorage.getItem(type + '-recent') || ''
    let json
    try {
      json = JSON.parse(jsonStr)
    } catch {
      json = []
    }

    return json
  }
  setRecentFiles(type: string, recent: any[]) {
    localStorage.setItem(type + '-recent', JSON.stringify(recent))
  }
  clearRecentFiles(type: string) {
    this.setRecentFiles(type, [])
  }
  addRecentFile(type: string, title: string, url: string, format: string) {
    const recent = this.getRecentFiles(type)

    let idx
    if ((idx = recent.findIndex(item => item.url === url)) > -1) {
      recent.splice(idx, 1)
    }
    recent.unshift({
      title,
      folder: Plugin.convPath(Path.dirname(url)),
      format,
      url,
    })

    this.setRecentFiles(type, recent)
  }
}

export const Store = new CStore()
