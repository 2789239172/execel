import utils from './utils.js'
import excelJs from "exceljs";
// import Worker from './excel.worker'

export default class Excel {
  constructor(file) {
    this.file = file
    // this.worker = new Worker()
    // this.worker.postMessage(file)
    this.workbook = new excelJs.Workbook()
    this.openSheet = null
    this.yMerges = null
  }

  async parseExcel() {
    try {
      const { result: arrayBuffer } = await this.fileReader()
      // 加载文件
      console.log("加载文件中...", arrayBuffer);
      await this.loadWorkbook(arrayBuffer)
      console.log("加载成功...");
      return this
    } catch (err) {
      console.log('加载excel失败: ' + err)
      return Promise.reject(err)
    }
  }

  // 加载文件
  fileReader() {
    const reader = new FileReader()
    const file = this.file
    return new Promise(resolve => {
      reader.onloadend = () => resolve(reader)
      reader.readAsArrayBuffer(file);
    })
  }

  // 加载工作簿
  loadWorkbook(arrayBuffer) {
    return this.workbook.xlsx.load(arrayBuffer)
  }

  // 加载工作表
  loadSheet(options = {sheetIndex: 0}) {
    const sheets = this.workbook.worksheets
    if (!sheets[options.sheetIndex]) return []

    console.log(`加载表格中${options.sheetIndex}...`);
    let sheet = this.workbook.worksheets[options.sheetIndex]
    let {sheetData, maxLength} = this.mapSheet(sheet)
    
    sheetData = this.fillSheet(sheetData, maxLength)

    this.loadMerges(sheet, sheetData)
    
    if (options.sheetIndex === 0) {
      this.openSheet = sheet
    }

    console.log("表格加载完成...", sheetData);
    return sheetData
  }

  // 生成表基础数据
  mapSheet(sheet) {
    const sheetData = []
    let maxLength = 0
    let formulaIndex
    let formulaFlag
    
    sheet.eachRow({ includeEmpty: true }, (row) => {
      const rowData = [];
      // 填入数据
      row.eachCell({ includeEmpty: true }, (cell, cellNumber) =>{
        if (cell.formula) {
          if (formulaFlag) return
          formulaFlag = true
          formulaIndex = cellNumber -1
          return
        }
        rowData.push({
          address: cell.address,
          value: cell.value || '',
          style: this.getStyle(cell.style),
          isMerge: false,
          rowspan: 1,
          colspan: 1,
        })
      }
      );

      sheetData.push(rowData);

      // 记录最长行
      if (rowData.length > maxLength) {
        maxLength = rowData.length;
      }
    });
    
    if (formulaFlag) {
      maxLength = formulaIndex
    }

    return {
      sheetData,
      maxLength
    }
  }

  // 填充工作表空白
  fillSheet(sheetData, maxLength) {
    console.log("填充表格中...");
    return sheetData.map((item, index) => {
      if (item.length > maxLength) {
        item.splice(maxLength, item.length - maxLength + 1)
        return item
      }
      let startNum = item.length;
      let fillLength = maxLength - startNum;
      let fillData = new Array(fillLength)
        .fill("")
        .map((item, i) => ({
          address: utils.numToLetter(startNum + i) + (index + 1),
          isMerge: false,
          rowspan: 1,
          colspan: 1,
          value: " ",
        }));
      return item.concat(fillData);
    });
  }

  // 加载合并工作表
  loadMerges(sheet, sheetData) {
    if (sheet.hasMerges) {
      console.log(sheet)
      console.log("添加单元格合并...");
      let yMerges = []

      // 记录每列合并的长度
      let xCount = new Array(sheetData.length).fill(0);
      for (const { $range } of Object.values(sheet._merges)) {
        let [[x1, x2], [y1, y2]] = utils.getRange($range);

        // 开始
        let xRange = x2 - x1 + 1;
        let yRange = y2 - y1 + 1;
        let startIndex = x1 - xCount[y1];

        let data = Object.assign(sheetData[y1][startIndex], {
          isMerge: true,
          yMerge: y1 !== y2,
          xMerge: x1 !== x2,
          colspan: xRange,
          rowspan: yRange,
        });

        for (let i = y1; i <= y2; i++) {
          sheetData[i].splice(x1 - xCount[i], xRange);
          // 这里注意减一, 因为依旧有一个占位的元素
          xCount[i] += xRange - 1;
        }
        
        sheetData[y1].splice(startIndex, 0, data);
        if (x1 === 0) {
          yMerges.push(...[y1, y2])
        }
      }

      this.yMerges = [...new Set(yMerges)].sort((a, b) => a - b)
    }

    return sheetData
  }

  // 获取表头的大概位置
  getHeaderIndex() {
    // if (this.yMerges.length == 1 && this.yMerges[0] == 0) {
    //   return 0
    // }
    let count = 0 
    for (const key of this.yMerges) {
      if (key !== count) break
      count++
    }
    // console.log('headerIndex', count)
    return count
  }

  getStyle(style) {
    let align = style.alignment || {}
    let fill = style.fill || {}
    let font = style.font || {}
    return {
      ['text-align']: align.horizontal,
      ['vertical-align']: align.vertical,
      ['background-color']: '#' + fill.fgColor?.argb?.slice(2),
      ['font-weight']: font.bold && 'bold',
      ['font-size']: font.size,
      ['font-family']: font.name,
      ['color']: '#' + font.color?.argb?.slice(2),
    }
  }

  // 下载
  async downLoad() {
    const buffer = await this.workbook.xlsx.writeBuffer();
    utils.downloadExcel(buffer);
  }

  async getBuffer() {
    const buffer = await this.workbook.xlsx.writeBuffer();
    return buffer
  }

  // 生成文件
  async getFile(name = 'excel') {
    const buffer = await this.workbook.xlsx.writeBuffer();
    const file = new File([buffer], name + ".xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    console.log(file)
    return file
  }
}