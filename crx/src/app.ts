import { BrowserApi, createBrowserApi } from './common/browser_api'
import { openDocument } from './editor/editor'
import { catchWindowError, log } from './utils/log'
import { getQuery, injectFavicon } from './utils/util'

async function main() {
  catchWindowError()
  const browserApi: BrowserApi = await createBrowserApi()
  const filePath = browserApi.getStreamInfo()?.originalUrl || getQuery('url') || ''

  const _url = location.origin + location.pathname
  log.debug('debug url:', `${_url}?url=${filePath}`)

  injectFavicon(filePath)
  openDocument(filePath)
}

main()
