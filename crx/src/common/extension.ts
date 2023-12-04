import { NP_ERROR_CODE } from '@/editor/plugin'

class CExtension {
  closeActiveTab() {
    chrome.tabs.query({ active: true }, tabs => {
      if (tabs.length > 0) {
        chrome.tabs.remove(tabs[0].id || 0)
      }
    })
  }
  openUrl(url: string) {
    chrome.tabs.query({}, tabs => {
      let found = false

      for (let i = 0; i < tabs.length; i++) {
        const tab = tabs[i]
        if (tab.url == url) {
          found = true
          this.activeTab(tab)
          break
        }
      }

      if (!found) {
        chrome.tabs.create({ url })
      }
    })
  }
  activeTab(tab: chrome.tabs.Tab) {
    chrome.tabs.update(tab.id, {
      active: true,
    })
  }
  async getTabTitleCount(title: string) {
    return new Promise(resolve => {
      const val = sessionStorage.getItem('title-count')
      if (val) {
        resolve(val)
        return
      }

      chrome.tabs.query({}, tabs => {
        const nums: Array<number> = []
        for (let i = 0; i < tabs.length; i++) {
          const tab = tabs[i]
          const matches = (tab.title || '').match(new RegExp(`(\\* )?${title}(\\d+)`))
          nums.push(Number(matches && matches[2]) || 0)
        }
        const count = Math.max(...nums) + 1
        sessionStorage.setItem('title-count', count.toString())

        resolve(count)
      })
    })
  }
  onceDisableDownloadHook() {
    const bg = chrome.extension.getBackgroundPage()
    bg && bg.onceDisableDownloadHook()
  }
  async sendToBackground(type: string, data: object, timeout = -1) {
    const pTimeOut = (timeout: number) => {
      return new Promise<x2tData>(resolve => {
        setTimeout(() => {
          resolve({
            error_code: NP_ERROR_CODE.timeout,
            error_msg: 'send to background timeout',
          })
        }, timeout)
      })
    }

    const pExecute = new Promise<x2tData>(resolve => {
      chrome.runtime.sendMessage({ type, data }, data => {
        resolve(data)
      })
    })

    const awaited = [pExecute]
    if (timeout > 0) {
      awaited.push(pTimeOut(timeout))
    }

    return Promise.race(awaited)
  }
}

export const Extension = new CExtension()
