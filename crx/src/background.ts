import { Extension } from './common/extension'
import { NpPlugin } from './editor/plugin'
import { catchWindowError } from './utils/log'
import { getDocumentType } from './utils/util'

const VIEW_URL = chrome.extension.getURL('app.html')

function main() {
  catchWindowError()
  bindEvents()
  listenMessage()
  // NpPlugin.installX2t().then(ret => {
  //   console.log('install_x2t ret: ', ret)
  // })
}

function bindEvents() {
  chrome.downloads.onDeterminingFilename.addListener(
    (
      downloadItem: chrome.downloads.DownloadItem,
      suggest: (suggestion?: chrome.downloads.DownloadFilenameSuggestion) => void
    ) => {
      suggest()

      const url = downloadItem.finalUrl || downloadItem.url
      if (!url) {
        return
      }

      let urlObj
      try {
        urlObj = new URL(url)
      } catch {
        return
      }

      const fileType = urlObj.pathname.split('.').pop() || ''
      const docType = getDocumentType(fileType)
      if (!docType) {
        return
      }
      if (downloadHookDisabled) {
        downloadHookDisabled = false
        return
      }

      chrome.downloads.cancel(downloadItem.id)
      const viewUrl = VIEW_URL + '?url=' + urlObj.href
      Extension.openUrl(viewUrl)
    }
  )
}
function listenMessage() {
  chrome.runtime.onMessage.addListener((message, sender, response) => {
    const { type, data } = message
    switch (type) {
      case 'x2t':
        NpPlugin.x2t(data.path, data.format, data.password).then(response)
        break
      case 'x2tBuf':
        NpPlugin.x2tBuf(data.isPdfPrint, data.base64Buf, data.path, data.format).then(response)
        break
      case 'fileSelect':
        NpPlugin.fileSelect(data.type).then(response)
        break
      case 'writeFile':
        NpPlugin.writeFile(data.base64Buf, data.relativePath).then(response)
        break
      case 'readFile':
        NpPlugin.readFile(data.path).then(response)
        break
      case 'setBinPath':
        NpPlugin.setBinPath(data.path).then(response)
        break
    }
    return true
  })
}

let downloadHookDisabled = false
window.onceDisableDownloadHook = () => {
  downloadHookDisabled = true
}

main()
