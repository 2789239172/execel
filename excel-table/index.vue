<template>
<!-- 
  ⭐ 标有此标签说明待完善
 -->
  <div class="excel-table">
    <!-- <div @click="down">xiazai</div> -->
    <div
      class="el-table el-table--fit el-table--border "
    >
      <table class="el-table__body" cellspacing="0" cellpadding="0" border="0">
        <tr class="el-table__row" v-for="(row, index) in headerData" :key="`h-${index}`">
          <th
            v-for="(cell, i) in row" :key="i"
            class="el-table_4_column_10"
            :class="{['y-merge']:cell.yMerge}"
            :style="cell.style"
            :rowspan="cell.rowspan"
            :colspan="cell.colspan"
          >
            <div
            class="clo-content"
            v-html="getValue(cell)"
            @input="inputChange($event, cell)"
            @blur="inputBlur(cell, index, i)"
            ></div>
            <div class="border"></div>
          </th>
        </tr>
        <!-- 表格主体 -->
        <tr class="el-table__row" v-for="(row, index) in pageData" :key="index">
          <td
            v-for="(cell, i) in row" :key="cell.address"
            class="el-table_4_column_10"
            :class="{['y-merge']:cell.yMerge}"
            :style="cell.style"
            :rowspan="cell.rowspan"
            :colspan="cell.colspan"
          >
            <el-select v-if="isSelect(i)" v-model="cell.value" @change="selectChange($event, cell)">
              <el-option
                v-for="(item, key) in isSelect(i)"
                :key="key"
                :label="item"
                :value="item">
              </el-option>
            </el-select>
             <div
             v-else
              contenteditable="true"
              class="clo-content"
              v-html="getValue(cell)"
              @input="inputChange($event, cell)"
              @blur="inputBlur(cell, index, i)"
            ></div>
            <div class="border"></div>
          </td>
        </tr>
      </table>
    </div>
  </div>
</template>

<script>
import Worker from './Excel/index.worker'
import util from './Excel/utils'
export default {
  name: "excel-table",
  props: {
    // 展示的表格
    uploadName: {
      type: String,
      default: ''
    },
    // 展示的表格
    sheetIndex: {
      type: Number,
      default: 0
    },

    // 是否可修改内容
    writable: {
      type: Boolean,
      default: true
    },

    // 每页展示行数 
    // 数据必须部分裁取, 否则页面会崩溃
    pageSize: {
      type: Number,
      default: 20
    },

    // 当前页数
    page: {
      type: Number,
      default: 0
    },

    fixedHeader: {
      type: Boolean,
      default: true
    },

    innerHeader: {
      type: Boolean,
      default: false
    }
  },
  data() {
    return {
      sheetData: [],
      headerData: [],

      selectData: {},
    };
  },
  watch: {
    // 切换表
    sheetIndex() {
      this.loadSheet()
    }
  },
  computed: {
    disabled() {
      return !this.writable
    },
    // 当前页的数据
    pageData() {
      let start = this.page * this.pageSize
      let end = start + this.pageSize

      return this.sheetData.slice(start, end) 
    },
  },
  mounted() {
    this.worker = new Worker()
    const execute = {
      imported: (excel) => {
        console.log('imported', excel)
        this.excel = excel
        this.loadSheet()
      },
      sheetLoaded: ({sheetData, headerIndex}) => {
        this.sheetData = sheetData
        console.log(this.sheetData)
        // 解析表头
        if (this.fixedHeader) {
          this.headerData = this.innerHeader ?
              this.sheetData.slice(headerIndex, headerIndex + 1) :
              this.sheetData.splice(0, headerIndex + 1)
        }
        console.log('this.headerData', this.headerData)
        // 返回长度
        const sheetLength = this.sheetData.length 
        this.$emit('sheetLoaded', {
          length: sheetLength,
          pageSize: this.pageSize,
          pageTotal: Math.ceil(sheetLength / this.pageSize)
        })
      },
      loadSheetListed: ({selectData}) => {
        console.log('selectData', selectData)
        this.selectData = selectData
      },
      getFile: (file) => {
        const formData = new FormData();
        formData.append("file", file);
        console.log(formData)
        this.$emit('getFormData', formData)
      },
      download({buffer}) {
        util.downloadExcel(buffer);
      }
    }
    this.worker.addEventListener("message", e => {
      const {key, value} = e.data
      execute[key](value)
    });
  },
  methods: {
    down() {
      this.worker.postMessage({
        key: 'download'
      })
    },

    // 导入
    async importExcel(file) {
      this.worker.postMessage({
        key: 'import',
        value: file
      })
    },

    // 解析sheet
    loadSheet() {
      this.worker.postMessage({
        key: 'loadSheet',
        value: this.sheetIndex
      })
      this.worker.postMessage({
        key: 'loadSheetList',
      })
    },

    // 判断下拉
    isSelect(index) {
      const header = this.headerData[this.headerData.length - 1]
      try {
        const selectKey = header[index].value
        const select = this.selectData[selectKey]
        return select
      }catch {
        return false
      }
    },

    // 修改表格
    inputChange({ target }, cell) {
      /* ⭐
        这里需要找到sheet对应元素修改value, 并且保证样式等不被修改
        暂时这里的样式会被修改
      */
      this.worker.postMessage({
        key: 'changeValue',
        value: {
          address: cell.address,
          value: target.innerText
        }
      })
      this.inputValue = target.innerText 
      this.inputFlag = true
    },

    selectChange(value, cell) {
      this.worker.postMessage({
        key: 'changeValue',
        value: {
          address: cell.address,
          value: value
        }
      })
    },

    // 录入数据  
    inputBlur(cell, index, i) {
      if (!this.inputFlag) return
      this.sheetData[index][i].value = this.inputValue
      this.inputFlag = false
    },

    // 解析单元格
    getValue(cell) {
      if (!cell?.value?.richText) return cell.value
      return cell.value.richText.reduce((old, now) => old + now.text, "");
    },

    // 导出文件
    downLoad() {
      this.excel?.downLoad()
    },

    // 获取文件
    getFile() {
      this.worker.postMessage({
        key: 'getFile'
      })
      // const fromData = new FormData();
      // const file = await this.excel.getFile()
      // fromData.append("file", file);
      // return fromData
    },

    // 上传
    async upload() {
      if (!this.excel) return
      const fromData = new FormData();
      const file = await this.excel.getFile(this.uploadName)

      fromData.append("file", file);

      this.$emit('upload', fromData)
    },
  },
};
</script>

<style scoped lang="scss">
[class*=-slot-wrapper] {
  display: inline-block;
}
.el-table {
  overflow-x: auto;
}
.el-table--border::after {
  width: 0 !important;
}

/* .excel-table {
  width: 100%;
  overflow-x: auto;
} */

td {
  padding: 0 !important;
  overflow: hidden;
  min-width: 100px !important;
  min-height: 50px;
  position: relative;
}

td:hover {
  background-color: #f5f7fa;
  transition: all 0.4s;
}

.clo-content {
  background: transparent;
  border: 0;
  width: 100%;
  height: 100%;
  padding: 10px;
  outline: 0;
  box-sizing: border-box;
  position: relative;
  z-index: 100;
}

.y-merge .clo-content{
  position: absolute;
  left: 0;
  right: 0;
  top: 0;
  bottom: 0;
}

textarea {
  resize: none;
}

table :focus + .border{
  position: absolute;
  left: 0;
  right: 0;
  top: 0;
  bottom: 0;
  border: 1px solid rgb(32, 238, 228);
}

/deep/ .el-select input {
  border: 0;
}
</style>
