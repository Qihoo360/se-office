import { c_oAscAsyncAction, c_oAscFileType2 } from './const'
import { Extension } from '../common/extension'
import { NP_ERROR_CODE, Plugin } from './plugin'
import { log } from '@/utils/log'
import { getDocumentType } from '@/utils/util'
import { saveAs } from 'file-saver'
import Path from 'path-browserify'
import { Store } from './store'

export let g_Editor: Editor

function initEditor(url: string, localPath: string, fileType: string, title: string) {
  return new Promise<DocsAPI.DocEditor>(resolve => {
    const docType = getDocumentType(fileType)
    const editor = new DocsAPI.DocEditor('iframe', {
      document: {
        title: title,
        url: url,
        fileType: fileType,
        permissions: {
          edit: true,
          chat: false,
          protect: false,
        },
      },
      editorConfig: {
        lang: 'zh',
        // targetApp: 'desktop',
        customization: {
          help: false,
          about: false,
          hideRightMenu: true,
          features: {
            spellcheck: {
              change: false,
            },
          },
          anonymous: {
            // set name for anonymous user
            request: false, // enable set name
            label: 'Guest', // postfix for user name
          },
        },
        templates: [
          {
            title: '文档',
            image: '../../common/main/resources/img/doc-formats/docx.svg',
            url: 'WORD',
          },
          {
            title: '表格',
            image: '../../common/main/resources/img/doc-formats/xlsx.svg',
            url: 'EXCEL',
          },
          {
            title: '演示文稿',
            image: '../../common/main/resources/img/doc-formats/pptx.svg',
            url: 'PPT',
          },
        ],
        recent: Store.getRecentFiles(docType),
      },
      events: {
        onAppReady: () => {
          resolve(editor)
        },
        onDocumentReady: () => {
          document.dispatchEvent(new CustomEvent('asc_onDocumentReady'))
        },
        onReopenDocument: (msg: MessageData) => {
          document.dispatchEvent(new CustomEvent('asc_onReopenDocument', { detail: msg.data }))
        },
        onRequestCreateNew: (msg: MessageData) => {
          let ext
          switch (msg.data.type) {
            case 'WORD':
              ext = '.docx'
              break
            case 'EXCEL':
              ext = '.xlsx'
              break
            case 'PPT':
              ext = '.pptx'
              break
          }
          chrome.tabs.create({ url: chrome.extension.getURL('app.html') + '?url=' + ext })
        },
        onRequestOpenFile: async (msg: MessageData) => {
          const data = await Plugin.fileSelect(msg.data.type)
          const path = data.data?.path
          if (path) {
            Extension.openUrl(path)
          }
        },
        onOpenRecent: (msg: MessageData) => {
          Extension.openUrl(msg.data.url)
        },
        onClearRecent: () => {
          Store.clearRecentFiles(docType)
        },
        onRequestSetDefault: (msg: MessageData) => {
        },
        onRequestFeedback: (msg: MessageData) => {
        },
        onAlertOK: (msg: MessageData) => {
          if (msg.data.button == 'custom') {
            location.reload()
          } else if (msg.data.type == 'closeWindow') {
            Extension.closeActiveTab()
          }
        },
        onDocumentCanSaveChanged: (msg: MessageData) => {
          chrome.mimeHandlerPrivate?.setShowBeforeUnloadDialog(msg.data.isCanSave)
        },
        onSave: async (msg: MessageData) => {
          log.debug('DocEditor.onSave', msg.data)

          const { data, option } = msg.data

          if (option.actionType == c_oAscAsyncAction.PrintDownload) {
            saveAs(option.downloadUrl, 'print.pdf')
            return
          }

          const { data: buf } = data
          const path = option.actionType == c_oAscAsyncAction.Save ? localPath : ''
          const ret = await Plugin.x2tBuf(
            !!option.isPdfPrint,
            buf,
            path,
            c_oAscFileType2[option.outputformat]
          )

          const saveCallbackData: { [key: string]: string | number } = {
            err_code: ret.error_code,
          }

          if (option.isPdfPrint) {
            let pdfUrl = ret.data?.path || '' // URL.createObjectURL(new Blob([docBuf], { type: 'application/pdf' }))
            if (!pdfUrl.startsWith('/')) {
              pdfUrl = '/' + pdfUrl
            }
            saveCallbackData['url'] = pdfUrl
          } else if (option.outputurls) {
            saveCallbackData['output.bin'] = ret.data?.path || ''
          } else if (
            option.actionType == c_oAscAsyncAction.Save ||
            option.actionType == c_oAscAsyncAction.DownloadAs
          ) {
            if (ret.data?.path) {
              if (localPath === '') {
                const newPath = 'file:///' + ret.data?.path
                localPath = Plugin.convPath(newPath)
                history.pushState({}, '', '?url=' + newPath)
                editor.sendCommand({
                  command: 'asc_setDocumentName',
                  data: localPath.split('\\').pop(),
                })
              }
            } else {
              switch (ret.error_code) {
                case NP_ERROR_CODE.none:
                case NP_ERROR_CODE.cancel:
                  break
                default:
                  EditorT(editor).alert('文件保存失败，请检查后重试。')
                  EditorT(editor).track('1551.950.gif', [ret.error_code, ret.error_msg])
                  break
              }
            }
          }

          editor.sendCommand({
            command: 'asc_onSaveCallback',
            data: saveCallbackData,
          })
        },
        onDownload: () => {
          Extension.onceDisableDownloadHook()

          const link = document.createElement('a')
          link.setAttribute('href', url)
          link.setAttribute('download', 'download')
          link.click()
        },
        writeFile: (msg: MessageData) => {
          const { file, data } = msg.data

          const reader = new FileReader()
          reader.onload = async e => {
            const base64 = e.target?.result as string
            const start = base64.indexOf(';base64,') + 8
            let path = Path.dirname(g_Editor.BinPath) + '/media/' + file
            if (!path.startsWith('/')) {
              path = '/' + path
            }

            const ret = await Plugin.writeFile(base64.substring(start), path)
            if (ret.error_code == 0) {
              const imageUrls: { [key: string]: string } = {}
              imageUrls[`media/${file}`] = path // URL.createObjectURL(new Blob([(data as any).buffer]))
              editor.sendCommand({
                command: 'asc_setImageUrls',
                data: {
                  urls: imageUrls,
                },
              })

              editor.sendCommand({
                command: 'asc_writeFileCallback',
                data: null,
              })
            } else {
              EditorT(editor).alert('写入文件失败，请检查后重试。')
            }
          }
          reader.readAsDataURL(new Blob([(data as any).buffer]))
        },
      },
    })
  })
}

function EditorT(editor: DocsAPI.DocEditor) {
  return new Editor(editor)
}
class Editor {
  private _editor: DocsAPI.DocEditor
  private _binPath = ''

  constructor(editor: DocsAPI.DocEditor) {
    this._editor = editor
    this.bindEvents()
  }

  static async create(url: string, localPath: string, fileType: string, title: string) {
    const editor = await initEditor(url, localPath, fileType, title)
    g_Editor = new Editor(editor)
    return g_Editor
  }

  set BinPath(v: string) {
    this._binPath = v
  }
  get BinPath() {
    return this._binPath
  }

  bindEvents() {
    const iframe = document.getElementsByName('frameEditor')[0] as HTMLIFrameElement
    iframe?.contentDocument?.addEventListener('click', e => {
      const srcEle = e.target as HTMLElement
      if (srcEle) {
        if (srcEle.getAttribute('download') !== null || srcEle.classList.contains('download')) {
          Extension.onceDisableDownloadHook()
        }
      }
    })
  }
  openDocument(buffer: Uint8Array) {
    this._editor.sendCommand({
      command: 'asc_openDocument',
      data: {
        buf: buffer,
      },
    })
  }
  openCSV(buffer: Uint8Array) {
    this._editor.sendCommand({
      command: 'asc_openEmptyDocument',
    })
    document.addEventListener('asc_onDocumentReady', () => {
      this._editor.sendCommand({
        command: 'asc_OpenCSV',
        data: {
          buf: buffer,
        },
      })
    })
  }
  setImageUrls(urls: MediaURLMap) {
    this._editor.sendCommand({
      command: 'asc_setImageUrls',
      data: {
        urls,
      },
    })
  }
  showPasswordDialog() {
    this._editor.sendCommand({
      command: 'asc_showPasswordDialog',
    })
  }
  alert(msg: string, type?: string) {
    this._editor.sendCommand({
      command: 'asc_Alert',
      data: {
        closable: false,
        iconCls: 'warn',
        title: '错误',
        msg,
        type: type,
      },
    })
  }
  alertRetry(msg: string, type?: string) {
    this._editor.sendCommand({
      command: 'asc_Alert',
      data: {
        closable: false,
        iconCls: 'warn',
        title: '错误',
        msg,
        type: type,
        buttons: ['custom', 'cancel'],
        customButtonText: '重试',
      },
    })
  }
  track(gif: string, _referer: string | Array<number | string | undefined>) {
    this._editor.sendCommand({
      command: 'asc_trackEvent',
      data: {
        gif,
        _referer,
      },
    })
  }
}

function renderDocument(editor: Editor, docData: DocData) {
  if (!(docData.data && docData.data.bin)) {
    return
  }

  g_Editor.BinPath = docData.data.path || ''
  if (g_Editor.BinPath === '') {
    const newPath = `plugin\\x2t\\out\\${Date.now()}\\new.bin`
    Plugin.setBinPath(newPath)
    g_Editor.BinPath = newPath.replace(/\\/g, '/')
  }

  if (docData.data.type == 'csv') {
    editor.openCSV(docData.data.bin)
  } else {
    editor.setImageUrls(docData.data.media)
    editor.openDocument(docData.data.bin)
  }
}

async function getDocTitle(fileName: string) {
  let title = ''
  fileName = fileName.toLocaleLowerCase()
  switch (fileName) {
    case '.docx':
      title = '文档' + (await Extension.getTabTitleCount('文档'))
      break
    case '.xlsx':
      title = '表格' + (await Extension.getTabTitleCount('表格'))
      break
    case '.pptx':
      title = '演示文稿' + (await Extension.getTabTitleCount('演示文稿'))
      break
  }

  return title
}
/**
 * 打开文档
 * 创建编辑器，文档转化bin
 * @param filePath
 */
export async function openDocument(url: string) {
  const online = !url.startsWith('file:///')

  const filePath = Plugin.convPath(url)
  const fileName = filePath.split(/[/\\]/).pop() || ''

  let fileType = fileName.split('.').pop() || ''
  fileType = fileType.toLocaleLowerCase()
  if (!getDocumentType(fileType)) {
    fileType = 'docx'
  }

  let localPath = filePath
  if (online) {
    localPath = ''
  }

  let title = await getDocTitle(fileName)
  if (!title) {
    title = fileName
    Store.addRecentFile(getDocumentType(fileType), title, url, fileType)
  } else {
    localPath = ''
  }

  document.title = title

  const [editor, docData] = await Promise.all([
    Editor.create(url, localPath, fileType, title),
    Plugin.getDocData(filePath),
  ])

  if (!(docData.data && docData.data.bin)) {
    switch (docData.error_code) {
      case NP_ERROR_CODE.need_password:
        editor.showPasswordDialog()

        document.addEventListener('asc_onReopenDocument', async (e: Event) => {
          const password = (e as CustomEvent).detail.password
          const docData = await Plugin.getDocData(url, password)
          renderDocument(editor, docData)
        })
        break
      default:
        {
          let suggest = ''
          if (online) {
            suggest = `请<a href="${url}" download><b class="download">下载到本地</b></a>后重试`
          } else {
            suggest = `请检查后重试`
          }
          if (docData.error_code === NP_ERROR_CODE.crash) {
            editor.alertRetry(`文件无法打开，${suggest}。`, 'closeWindow')
          } else {
            editor.alert(`文件无法打开，${suggest}。`, 'closeWindow')
          }
          editor.track('1551.8744.gif', [docData.error_code, docData.error_msg])
        }
        break
    }
    return
  }

  renderDocument(editor, docData)
}
