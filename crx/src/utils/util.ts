export function getDocumentType(fileType: string | undefined) {
  fileType = fileType?.toLowerCase() ?? ''
  const type =
    /^(?:(xls|xlsx|ods|csv|gsheet|xlsm|xlt|xltm|xltx|fods|ots|xlsb|sxc|et|ett)|(pps|ppsx|ppt|pptx|odp|gslides|pot|potm|potx|ppsm|pptm|fodp|otp|sxi|dps|dpt)|(doc|docx|odt|gdoc|rtf|epub|djvu|xps|oxps|docm|dot|dotm|dotx|fodt|ott|fb2|oform|docxf|sxw|stw|wps|wpt))$/.exec(
      fileType
    )

  let documentType = ''
  if (!type) {
    documentType = ''
  } else {
    if (typeof type[1] === 'string') documentType = 'excel'
    else if (typeof type[2] === 'string') documentType = 'ppt'
    else if (typeof type[3] === 'string') documentType = 'word'
  }

  return documentType
}

export function injectFavicon(path: string) {
  const fileType = path.split('.').pop() || ''
  const docType = getDocumentType(fileType)

  if (docType) {
    document.head.insertAdjacentHTML(
      'beforeend',
      `<link rel="shortcut icon" type="image/ico" href="${chrome.extension.getURL(
        'img/icon/' + docType + '.ico'
      )}">`
    )
  }
}

export function getQuery(key: string) {
  let val = new URL(location.href).search.replace(`?${key}=`, '')
  try {
    val = decodeURI(val)
  } catch {
    //
  }

  return val
}

export async function getFileSize(path: string) {
  const req = await fetch(path).catch(() => {})
  if (!req) {
    return -1
  }

  const blob = await req.blob()
  return blob.size
}
