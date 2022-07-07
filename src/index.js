import './app.css'
import Excel from 'exceljs'

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

  let modalId = 'myModal' + fileInfo.fileKey
  let tabId = 'myTab' + fileInfo.fileKey
  let tabContentId = 'tab-content' + fileInfo.fileKey
  let $button = $(
    '<button type="button" class="btn btn-default" data-toggle="modal" data-target="#' +
      modalId +
      '"><span class="fa fa-search"></span></button>',
  )

  let myModal =
    '<style type="text/css">td{word-break: keep-all;white-space:nowrap;}</style>' +
    '<div class="modal fade tab-pane active" id="' +
    modalId +
    '" tabindex="-1" role="dialog" aria-labelledby="myModalLabel">' +
    '<div class="modal-dialog modal-xl" style="border-radius:5px" role="document">' +
    '<div class="modal-content">' +
    '<div class="modal-header">' +
    '<h5 class="modal-title">' +
    fileInfo.name +
    '</h5>' +
    '<button type="button" class="close" data-dismiss="modal" aria-label="Close">' +
    '<span aria-hidden="true">&times;</span>' +
    '</button>' +
    '</div>' +
    '<ul class="nav nav-tabs" id=' +
    tabId +
    '>' +
    '</ul>' +
    '<div id=' +
    tabContentId +
    ' class="tab-content">' +
    '<div class="d-flex justify-content-center">' +
    '<div class="spinner-border" role="status">' +
    '<span class="sr-only">Loading...</span>' +
    '</div>' +
    '</div>' +
    '</div>' +
    '<div class="modal-footer">' +
    '<button type="button" class="btn btn-secondary" data-dismiss="modal">关闭</button>' +
    '</div>' +
    '</div>' +
    '</div>' +
    '</div>'

  previewElement.append($button)

  $('body').prepend(myModal)
  $('#' + modalId).on('shown.bs.modal', function (e) {
    loadRemoteFile(fileInfo)
  })
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
      const workbook = new Excel.Workbook()
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
        const blob = new Blob([uint8Array], { type: 'application/octet-binary' })
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
