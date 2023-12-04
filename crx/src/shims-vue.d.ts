/* eslint-disable */

declare module '*.vue' {
  import type { DefineComponent } from 'vue'
  const component: DefineComponent<{}, {}, any>
  export default component
}

declare module 'min-log'

declare interface Window {
  onceDisableDownloadHook: () => void
}

declare namespace DocsAPI {
  export class DocEditor {
    constructor(string, any)

    sendCommand(message: {})
  }
}

interface MediaURLMap {
  [key: string]: string
}

interface x2tData {
  error_code: number
  error_msg: string
  data?: {
    path: string
    media: Array<string>
  }
}
interface DocData {
  error_code?: number
  error_msg?: string
  data?: {
    type: string | undefined
    bin: Uint8Array | void
    media: MediaURLMap
    path?: string
  }
}

interface MessageData {
  target: DocEditor
  data: {
    option: {
      isPdfPrint: boolean
      nobase64: boolean
      outputformat: number
      title: string
      actionType: number
      downloadUrl: string
      outputurls: boolean
    }
    file: string
    data: {
      data: string | Uint8Array
    }
    type: string
    isCanSave: boolean
    button: string
    url: string
  }
}

interface Plugin {
  [key: string]: (...params) => void
  x2t(path: string, format: string, callback: string)
  x2t_buf(
    is_pdf_print: boolean,
    base64_buf: string | Uint8Array,
    path: string,
    format: string,
    callback: string
  )
  file_select(type: string, callback: string)
  write_file(base64_buf: string | Uint8Array, relative_path: string, callback: string)
  install_x2t(callback: string)
}
declare var plugin: Plugin
