import ExcelJS from 'exceljs'

const checkXls = (file) => {
  let reg = /\.xl(s[xmb]|t[xm]|am|s)$/g
  return reg.test(file)
}

const loadModal = (fileInfo) => {
  let previewElement
  jQuery('.file-image-container-gaia').each(function (i, e) {
    let fileName = jQuery(e).children('a:eq(0)').text()
    if (fileName == fileInfo.name && jQuery(e).children('button').length == 0) {
      previewElement = jQuery(e)
      return false
    }
  })

  if (!previewElement) return

  let $span = $('<a href="javascript:;">下载</a>')
  $span.click(() => {
    loadRemoteFile(fileInfo)
  })
  previewElement.append($span)
}

const loadRemoteFile = (fileInfo) => {
  let fileUrl = '/k/v1/file.json?fileKey=' + fileInfo.fileKey
  readWorkbookFromRemoteFile(fileUrl)
}

const readWorkbookFromRemoteFile = async (url) => {
  let xhr = new XMLHttpRequest()
  xhr.open('get', url, true)

  xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest')
  xhr.responseType = 'arraybuffer'
  xhr.onload = function (e) {
    if (xhr.status == 200) {
      let data = new Uint8Array(xhr.response)
      const workbook = new ExcelJS.Workbook()
      workbook.xlsx.load(data).then(async (resp) => {
        console.log(resp)

        workbook.addWorksheet('demo')
        const arraySheet = workbook.getWorksheet('demo')

        arraySheet.columns = [
          { header: 'ID', key: 'id' },
          { header: '姓名', key: 'name' },
          { header: '年龄', key: 'age' },
        ]
        arraySheet.addRows(dummyData)

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
      })
    }
  }
  xhr.send()
}

const dummyData = [...Array(5)].map((_, i) => {
  return {
    id: i,
    name: 'name' + i,
    age: Math.floor(Math.random() * 20) + 20,
  }
})

kintone.events.on('app.record.detail.show', function (event) {
  let record = event.record
  for (let index in record) {
    let field = record[index]
    if (field.type === 'FILE') {
      let fieldValue = field.value
      fieldValue.forEach(function (file) {
        if (checkXls(file.name)) {
          loadModal(file)
        }
      })
    }
  }
})
