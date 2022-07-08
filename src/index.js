import ExcelJS from 'exceljs'

const templateField = 'template'

const checkXls = (file) => {
  const reg = /\.xl(s[xmb]|t[xm]|am|s)$/g
  return reg.test(file)
}

const loadDownload = (fileInfo, record) => {
  if (document.getElementById('downloadButton') !== null) {
    return
  }
  const downloadButton = document.createElement('button')
  downloadButton.id = 'downloadButton'
  downloadButton.innerText = '一键生成excel'
  downloadButton.onclick = () => {
    const fileUrl = '/k/v1/file.json?fileKey=' + fileInfo.fileKey
    readWorkbookFromRemoteFile(fileUrl, record)
  }
  kintone.app.record.getHeaderMenuSpaceElement().appendChild(downloadButton)
}

const readWorkbookFromRemoteFile = async (url, record) => {
  const xhr = new XMLHttpRequest()
  xhr.open('get', url, true)

  xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest')
  xhr.responseType = 'arraybuffer'
  xhr.onload = async (e) => {
    if (xhr.status == 200) {
      let data = new Uint8Array(xhr.response)
      const workbook = new ExcelJS.Workbook()
      await workbook.xlsx.load(data)
      const arraySheet = workbook.getWorksheet('demo')
      const kintoneData = [record.name.value, record.model.value, record.price.value]
      const index = 3

      arraySheet.addRow(kintoneData, index)

      const uint8Array = await workbook.xlsx.writeBuffer()
      const blob = new Blob([uint8Array], {
        type: 'application/octet-binary',
      })
      const a = document.createElement('a')
      const url = window.URL.createObjectURL(blob)
      a.href = url
      a.download = 'new.xlsx'
      a.click()
      window.URL.revokeObjectURL(url)
    }
  }
  xhr.send()
}

kintone.events.on('app.record.detail.show', (event) => {
  const record = event.record
  const file = record[templateField].value
  if (file.length > 0) {
    if (checkXls(file[0].name)) {
      loadDownload(file[0], record)
    }
  }
  return event
})
