import { log } from '@/utils/log'
import { emptyTypeKeyMap, g_sEmpty_bin } from './empty_bin'
import { Extension } from '@/common/extension'
import { getDocumentType, getFileSize } from '@/utils/util'

export const NP_ERROR_CODE = {
  none: 0,
  // system
  cancel: -4,
  catch: -11,
  crash: -12,
  timeout: -13,
  // x2t
  need_password: 101,
}

class CNpPlugin {
  async x2t(path: string, format: string, password?: string) {
    return new Promise<x2tData>(resolve => {
      const callbackName = this._createCallback(resolve, 'x2t')
      log.debug(
        `before plugin.x2t('${path}', '${format}', '${callbackName}')`,
        Date.now() - performance.timing.navigationStart
      )
      // plugin.x2t(path, format, callbackName)
      this.callNP('x2t', path, format, callbackName)
    })
  }
  async x2tBuf(isPdfPrint: boolean, base64Buf: string | Uint8Array, path: string, format: string) {
    return new Promise<x2tData>(resolve => {
      const callbackName = this._createCallback(resolve, 'x2t_buf')
      log.debug(
        `before plugin.x2t_buf(${isPdfPrint}, *base64*, '${path}', '${format}', '${callbackName}')`
      )
      this.callNP('x2t_buf', isPdfPrint, base64Buf, path, format, callbackName)
    })
  }
  async fileSelect(type: string) {
    const MAP: { [key: string]: string } = {
      word: '.docx',
      excel: '.xlsx',
      ppt: '.pptx',
    }
    type = MAP[type] || type

    return new Promise<x2tData>(resolve => {
      const callbackName = this._createCallback(resolve, 'file_select')
      log.debug(`before plugin.file_select('${type}', '${callbackName}')`)
      // plugin.file_select(type, callbackName)
      this.callNP('file_select', type, callbackName)
    })
  }
  async writeFile(base64Buf: string | Uint8Array, relativePath: string) {
    return new Promise<x2tData>(resolve => {
      const callbackName = this._createCallback(resolve, 'write_file')
      log.debug(`before plugin.write_file(*base64*, '${relativePath}', '${callbackName}')`)
      this.callNP('write_file', base64Buf, relativePath, callbackName)
    })
  }
  async readFile(path: string) {
    return new Promise<x2tData>(resolve => {
      const callbackName = this._createCallback(resolve, 'read_file')
      log.debug(`before plugin.read_file('${path}', '${callbackName}')`)
      this.callNP('read_file', path, callbackName)
    })
  }
  async setBinPath(path: string) {
    return new Promise<x2tData>(resolve => {
      const callbackName = this._createCallback(resolve, 'set_bin_path')
      log.debug(`before plugin.set_bin_path('${path}', '${callbackName}')`)
      this.callNP('set_bin_path', path, callbackName)
    })
  }
  async installX2t() {
    return new Promise<x2tData>(resolve => {
      const callbackName = this._createCallback(resolve, 'install_x2t')
      log.debug(`before plugin.install_x2t('${callbackName}')`)
      plugin.install_x2t(callbackName)
      // this.callNP('install_x2t', callbackName)
    })
  }
  callNP(method: string, ...params: any[]) {
    const func = plugin[method]
    if (func) {
      func.apply(plugin, params)
    } else {
      const callback = window[params.pop()] as any
      callback &&
        callback(`
          {
            "error_code": ${NP_ERROR_CODE.crash},
            "error_msg": "NP process crash"
          }
        `)
      window.requestIdleCallback(() => {
        location.reload()
      })
    }
  }

  async getDocData(path: string, password?: string): Promise<DocData> {
    const ext = path.split('.').pop()?.toLocaleLowerCase()

    const fileSize = await getFileSize(path)
    if (fileSize === 0) {
      const docType = getDocumentType(ext)
      path = emptyTypeKeyMap[docType]
    }

    const emptyBin = g_sEmpty_bin[path]
    if (emptyBin) {
      return {
        data: {
          type: ext,
          bin: emptyBin as any,
          media: {},
        },
      }
    }

    if (ext == 'csv') {
      let error_msg = ''
      const req = await fetch(path).catch(err => {
        error_msg = err.message + ': ' + path
      })
      if (!req) {
        return {
          error_msg,
        }
      }
      const res = await req.arrayBuffer()
      return {
        data: {
          type: ext,
          bin: new Uint8Array(res),
          media: {},
        },
      }
    }

    const start = Date.now()
    const ret = await this.x2t(path, 'bin', password)
    log.debug('plugin.x2t() 耗时：', Date.now() - start, ret)

    if (!(ret.data && ret.data.path)) {
      return ret as DocData
    }

    let error = null
    const buf = await this.fetchBin(encodeURI(ret.data.path)).catch(err => {
      error = {
        error_msg: err.message,
      }
    })
    if (error) {
      return error
    }

    const media: MediaURLMap = {}
    ;(ret.data.media || []).forEach((item: string) => {
      const name = item.split('/').pop()
      const imgUrl = `media/${name}`
      media[imgUrl] = '/' + item
    })

    return {
      data: {
        type: ext,
        bin: buf,
        media: media,
        path: ret.data.path,
      },
    }
  }
  async fetchBin(url: string) {
    const req = await fetch(url)
    const res = await req.arrayBuffer()
    return new Uint8Array(res)
  }

  private _createCallback(resolve: (value: x2tData) => void, funcName: string) {
    const callbackName = `_x2t_callback_${Date.now()}`
    ;(window as any)[callbackName] = (ret: string) => {
      ret = ret.replace(/\\/g, '/')
      let data = null
      try {
        data = JSON.parse(ret) as x2tData
      } catch (ex) {
        data = {
          error_code: NP_ERROR_CODE.catch,
          error_msg: (ex as Error).message,
        }
      }
      log.debug(`end plugin.${funcName}()`, data)
      log.verbose(`plugin.${funcName}() return:`, ret)

      resolve(data)
    }

    return callbackName
  }
  convPath(path: string) {
    try {
      path = decodeURI(path)
    } catch { }

    if (path.startsWith('file:///')) {
      path = path.replace('file:///', '')
      path = path.replace(/\//g, '\\')
    }

    return path
  }
}

class CBgPlugin extends CNpPlugin {
  override async setBinPath(path: string) {
    return Extension.sendToBackground('setBinPath', { path })
  }
  override async x2t(path: string, format: string, password?: string | undefined) {
    return Extension.sendToBackground('x2t', { path, format, password })
  }
  override async x2tBuf(
    isPdfPrint: boolean,
    base64Buf: string | Uint8Array,
    path: string,
    format: string
  ) {
    if (format) {
      format = format.toLocaleLowerCase()
    }
    return Extension.sendToBackground('x2tBuf', { isPdfPrint, base64Buf, path, format })
  }
  override async fileSelect(type: string) {
    return Extension.sendToBackground('fileSelect', { type })
  }
  override async writeFile(base64Buf: string | Uint8Array, relativePath: string) {
    return Extension.sendToBackground('writeFile', { base64Buf, relativePath })
  }
}

export const NpPlugin = new CNpPlugin()
export const Plugin = new CBgPlugin()
