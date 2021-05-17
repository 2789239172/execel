import Excel from "./index";
let excel

const execute = {
  // 导入文件
  async import(file) {
    console.log('import')

    excel = new Excel(file)
    await excel.parseExcel()
    postMessage({
      key: 'imported',
      value: excel
    })
  },

  // 加载表
  loadSheet(index) {
    const sheetData = excel.loadSheet({sheetIndex: index})
    console.log('sheetLoaded', sheetData)
    postMessage({
      key: 'sheetLoaded',
      value: {
        sheetData,
        headerIndex: excel.getHeaderIndex()
      }
    })
  },

  loadSheetList() {
    const sheetData = excel.loadSheet({sheetIndex: 1})
    const sheetTitle = sheetData.shift()
    if (!sheetTitle) return
    const selectData = {}
    sheetTitle.forEach((item, index) => {
      const data = [...new Set(sheetData.map(item => item[index].value))]
      
      for (let i = data.length - 1; i >= 0; i--) {
        const item = data[i];
        if (item.trim && item.trim() == '') {
          data.pop()
        } else {
          break
        }
      }
      selectData[item.value] = data
    });
    postMessage({
      key: 'loadSheetListed',
      value: {
        selectData,
      }
    })
  },

  // 修改内容
  changeValue({address, value}) {
    const openCell = excel.openSheet.getCell(address);
    openCell.value = value;
  },

  // 拿取文件
  async getFile() {
    const file = await excel.getFile()

    postMessage({
      key: 'getFile',
      value: file
    })
  }
}
onmessage = (e) => {
  const {key, value} = e.data
  execute[key](value)
}