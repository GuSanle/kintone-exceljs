import ExcelJS from 'exceljs'

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
      const blob = await FillData(xhr, record)
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

const FillData = async (xhr, record) => {
  let data = new Uint8Array(xhr.response)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(data)
  const worksheet = workbook.getWorksheet('Sheet1')

  const CustomerNameCell = worksheet.getCell('C4')
  const AddressCell = worksheet.getCell('C7')
  const ContactCell = worksheet.getCell('C6')
  const DepartmentCell = worksheet.getCell('C5')
  const NoteCell = worksheet.getCell('C8')
  const PhoneCell = worksheet.getCell('H6')
  const ReceiveDateCell = worksheet.getCell('H5')
  const SnCell = worksheet.getCell('H4')
  const CreatorCell = worksheet.getCell('B18')
  CustomerNameCell.value = record.CustomerName.value
  AddressCell.value = record.Address.value
  ContactCell.value = record.Contact.value
  DepartmentCell.value = record.Department.value
  NoteCell.value = record.Note.value
  PhoneCell.value = record.Phone.value
  ReceiveDateCell.value = record.ReceiveDate.value
  SnCell.value = record.Sn.value
  CreatorCell.value = record.Creator.value
  let indexLine = 11
  for (const row of record.Table.value) {
    const rowValue = row.value
    const rowMap = {
      0: 'A',
      1: 'D',
      2: 'F',
      3: 'H',
    }
    Object.keys(rowValue).forEach((key, index) => {
      const cellField = rowMap[index] + indexLine
      const cell = worksheet.getCell(cellField)
      cell.value = rowValue[key].value
    })
    indexLine++
  }
  const uint8Array = await workbook.xlsx.writeBuffer()
  const blob = new Blob([uint8Array], {
    type: 'application/octet-binary',
  })
  return blob
}

kintone.events.on('app.record.detail.show', (event) => {
  const record = event.record
  const file = record.Template.value
  if (file.length > 0) {
    if (checkXls(file[0].name)) {
      loadDownload(file[0], record)
    }
  }
  return event
})
