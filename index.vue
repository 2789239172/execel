<template>
   <div class="excel-upload upload">
     <!-- 
        :on-remove="handleRemove"
        :on-exceed="handleExceed"
        :http-request="uploadFile"
      -->
      <el-upload
        drag
        action
        element-loading-text="加载文件中"
        v-loading="loading"
        v-show="showImport"
        class="upload-demo"
        ref="upload"
        accept=".xlsx"
        :file-list="fileList"
        :on-change="handleChange"
        :on-success="handleSuccess"
        :on-error="handleError"
        :before-upload="beforeUpload"
        :multiple="false"
        :show-file-list="false"
        :auto-upload="false"
      >
        <div class="el-upload__text" slot="trigger">
          <img
            src="~@/assets/image/excel.png"
            alt="文件上传"
          />
        </div>
      </el-upload>
      
      <div v-show="showDesc" class="desc">只能上传EXCLE文件</div>

      <div v-show="showTable">
        <div class="button-wrapper">
          <el-button  size="small" type="success" :loading="uploadLoading" @click="submitUpload">上传到服务器</el-button>
          <el-upload
            style="margin-left: 10px"
            drag
            action
            v-loading="loading"
            class="upload-reload"
            accept=".xlsx"
            :on-change="handleChange"
            :multiple="false"
            :show-file-list="false"
            :auto-upload="false"
          >
            <div class="el-button el-button--warning el-button--small" slot="trigger">
              重新导入
            </div>
          </el-upload>
        </div>

        <!-- 表格 -->
        <excel-table 
          ref="table"
          :page="tablePage"
          :page-size="pageSize"
          @sheetLoaded="sheetLoaded"
          @getFormData="getFormData"  
        ></excel-table>

        <!-- 分页 -->
        <el-pagination
          v-if="pageTotal"
          background
          layout="total, sizes, prev, pager, next, jumper"
          :page-size.sync='pageSize'
          :total="sizeTotal"
          :page-count="pageTotal"
          @current-change="pagneChange"
        ></el-pagination>
      </div>
   </div>
</template>

<script>
import excelTable from './excel-table'
export default {
   name: 'excel-upload',
   components: {excelTable},
   data() {
       return {
        fileList: [],
        loading: false,
        uploadLoading: false,
        showImport: true,

        // 分页
        pageTotal: 0,
        pageSize: 10,
        sizeTotal: 0,
        tablePage: 0
       };
   },
   computed: {
     showTable() {
       return !this.showImport && !this.loading
     },
     showDesc() {
       return !this.fileList.length
     },
     showSubmit() {
       return this.fileList.length && !this.loading
     }
   },
   methods: {
    // 表格加载完成
    sheetLoaded({pageTotal, length}) {
        this.pageTotal = pageTotal
        this.sizeTotal = length
        this.$nextTick(() => {
          this.showImport = false
          this.loading = false
        }) 
    },

    pagneChange(current) {
      this.tablePage = current - 1
    },

    submitUpload() {
      this.uploadLoading = true
      this.$refs.table.getFile()
    },

    getFormData(formData) {
      this.uploadLoading = false
      this.$emit('upload', formData)
    },

    // 上传文件之前的钩子, 参数为上传的文件,若返回 false 或者返回 Promise 且被 reject，则停止上传
    beforeUpload(file) {
      let extension = file.name.substring(file.name.lastIndexOf(".") + 1);
      let size = file.size / 1024 / 1024;
      if (extension !== "xlsx") {
        this.$message.warning("只能上传后缀是.xlsx的文件");
      }
      if (size > 10) {
        this.$message.warning("文件大小不得超过10M");
      }
    },

     // 文件上传成功
    handleSuccess(res, file, fileList) {
      this.$message.success("文件上传成功");
    },

    // 文件上传失败
    handleError(err, file, fileList) {
      this.$message.error("文件上传失败");
    },

    handleChange(file) {
      console.log(file)
      this.showImport = true
      this.loading = true
      this.fileList = [file]
      this.$refs.table.importExcel(file.raw)
    },
   },
}
</script>

<style scoped lang="scss">
.upload-demo {
  display: inline-block;
}
.upload-reload {
  display: flex;
  align-items: center;
}
.button-wrapper {
  display: flex;
  align-items: center;
  margin-bottom: 10px
}
/deep/ .el-upload-dragger {
  width: auto !important;
  height: auto !important;
  border: 0 !important;
}
</style>
