export const c_oAscAsyncAction = {
  Open: 0, // открытие документа
  Save: 1, // сохранение
  LoadDocumentFonts: 2, // загружаем фонты документа (сразу после открытия)
  LoadDocumentImages: 3, // загружаем картинки документа (сразу после загрузки шрифтов)
  LoadFont: 4, // подгрузка нужного шрифта
  LoadImage: 5, // подгрузка картинки
  DownloadAs: 6, // cкачать
  Print: 7, // конвертация в PDF и сохранение у пользователя
  UploadImage: 8, // загрузка картинки

  ApplyChanges: 9, // применение изменений от другого пользователя.

  SlowOperation: 11, // медленная операция
  LoadTheme: 12, // загрузка темы
  MailMergeLoadFile: 13, // загрузка файла для mail merge
  DownloadMerge: 14, // cкачать файл с mail merge
  SendMailMerge: 15, // рассылка mail merge по почте
  ForceSaveButton: 16,
  ForceSaveTimeout: 17,
  Waiting: 18,
  Submit: 19,
  Disconnect: 20,

  PrintDownload: 101,
}

//files type for Saving & DownloadAs
export const c_oAscFileType = {
  UNKNOWN: 0,
  PDF: 0x0201,
  PDFA: 0x0209,
  DJVU: 0x0203,
  XPS: 0x0204,

  // Word
  DOCX: 0x0041,
  DOC: 0x0042,
  ODT: 0x0043,
  RTF: 0x0044,
  TXT: 0x0045,
  HTML: 0x0046,
  MHT: 0x0047,
  EPUB: 0x0048,
  FB2: 0x0049,
  MOBI: 0x004a,
  DOCM: 0x004b,
  DOTX: 0x004c,
  DOTM: 0x004d,
  FODT: 0x004e,
  OTT: 0x004f,
  DOC_FLAT: 0x0050,
  DOCX_FLAT: 0x0051,
  HTML_IN_CONTAINER: 0x0052,
  DOCX_PACKAGE: 0x0054,
  OFORM: 0x0055,
  DOCXF: 0x0056,
  DOCY: 0x1001,
  CANVAS_WORD: 0x2001,
  JSON: 0x0808, // Для mail-merge

  // Excel
  XLSX: 0x0101,
  XLS: 0x0102,
  ODS: 0x0103,
  CSV: 0x0104,
  XLSM: 0x0105,
  XLTX: 0x0106,
  XLTM: 0x0107,
  XLSB: 0x0108,
  FODS: 0x0109,
  OTS: 0x010a,
  XLSX_FLAT: 0x010b,
  XLSX_PACKAGE: 0x010c,
  XLSY: 0x1002,

  // PowerPoint
  PPTX: 0x0081,
  PPT: 0x0082,
  ODP: 0x0083,
  PPSX: 0x0084,
  PPTM: 0x0085,
  PPSM: 0x0086,
  POTX: 0x0087,
  POTM: 0x0088,
  FODP: 0x0089,
  OTP: 0x008a,
  PPTX_PACKAGE: 0x008b,

  //image
  IMG: 0x0400,
  JPG: 0x0401,
  TIFF: 0x0402,
  TGA: 0x0403,
  GIF: 0x0404,
  PNG: 0x0405,
  EMF: 0x0406,
  WMF: 0x0407,
  BMP: 0x0408,
  CR2: 0x0409,
  PCX: 0x040a,
  RAS: 0x040b,
  PSD: 0x040c,
  ICO: 0x040d,
}

export const c_oAscFileType2 = Object.fromEntries(
  Object.entries(c_oAscFileType).map(([key, value]) => [value, key])
)
