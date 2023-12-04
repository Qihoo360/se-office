import minLog from 'min-log'

let _logLevel
if ((_logLevel = localStorage.getItem('__log-level'))) {
  minLog.setLevel(_logLevel)
} else if (document.referrer.includes('/dev/') || window.chrome) {
  minLog.setLevel('debug')
}

export const log = minLog.getLogger('CRX')

export function catchWindowError() {
}

function queryParams(source: { [key: string]: any }) {
  const array = []

  for (const key in source) {
    array.push(encodeURIComponent(key) + '=' + encodeURIComponent(source[key]))
  }

  return array.join('&')
}

const Analytics = (function () {
  return {
    basePath: '',

    send: function (gif: string, _referer: string | [] = '') {
      if (Object.prototype.toString.call(_referer) === '[object Array]') {
        _referer = (_referer as []).join('|')
      }
      const r = Date.now() + Math.random().toString().replace('0.', '').substr(0, 10)
      const params = {
        _referer,
        r,
      }

      const url = this.basePath + gif + '?' + queryParams(params)
      new Image().src = url
    },
  }
})()
