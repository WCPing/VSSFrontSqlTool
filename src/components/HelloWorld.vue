<template>
  <div class="index" v-loading.fullscreen.lock="fullscreenLoading" element-loading-text="拼命加载中...">
    <input type="file" @change="importFile(this)" id="imFile" style="display: none"
           accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"/>
    <a id="downlink"></a>
    <el-button class="button" @click="uploadFile()">导入</el-button>
    <el-button class="button" @click="downloadFile(excelData)">导出成excel</el-button>
    <el-select v-model="selType" placeholder="请选择sql类别">
      <el-option
        v-for="item in options" :key="item" :label="item" :value="item">
      </el-option>
    </el-select>
    <el-button class="button" @click="exportSQLFile()">导出成sql</el-button>
    <el-button type="danger" class="button" @click="sqlArr = []">清空</el-button>
    <!--错误信息提示-->
    <el-dialog title="提示" v-model="errorDialog" size="tiny">
      <span>{{errorMsg}}</span>
        <span slot="footer" class="dialog-footer">
          <el-button type="primary" @click="errorDialog=false">确认</el-button>
        </span>
    </el-dialog>
    <!--展示导入信息-->
    <el-table :data="excelData" stripe tooltip-effect="dark">
      <el-table-column label="KETID" prop="keyid" show-overflow-tooltip></el-table-column>
      <el-table-column label="ZH" prop="ZH" show-overflow-tooltip></el-table-column>
      <el-table-column label="EN" prop="EN" show-overflow-tooltip></el-table-column>
    </el-table>
    <div v-for="arr of sqlArr" :key="arr" class="sqlCss">
      {{arr}}
    </div>
  </div>
</template>

<script>
  // 引入xlsx
  var XLSX = require('xlsx')
  import _ from 'lodash'
  export default {
    name: 'Index',
    data () {
      return {
        fullscreenLoading: false, // 加载中
        imFile: '', // 导入文件el
        outFile: '',  // 导出文件el
        errorDialog: false, // 错误信息弹窗
        errorMsg: '', // 错误信息内容
        excelData: [
            {
              keyid: '页面标题/feedbackList', ZH: '问题反馈查看', EN: 'Feedback View'
            }
          ],
        options: ['资料','采购','无租户','物流','提示信息','质管','系统','账款','首页','页面标题','公共','表单验证'],
        selType: '',
        sqlArr: []
      }
    },
    mounted () {
      this.imFile = document.getElementById('imFile')
      this.outFile = document.getElementById('downlink')
    },
    methods: {
      uploadFile: function () { // 点击导入按钮
        this.imFile.click()
      },
      downloadFile: function (rs) { // 点击导出按钮
        let data = [{}]
        for (let k in rs[0]) {
          data[0][k] = k
        }
        data = data.concat(rs)
        this.downloadExl(data, '界面测试')
      },
      importFile: function () { // 导入excel
        this.fullscreenLoading = true
        let obj = this.imFile
        if (!obj.files) {
          this.fullscreenLoading = false
          return
        }
        var f = obj.files[0]
        var reader = new FileReader()
        let $t = this
        reader.onload = function (e) {
          var data = e.target.result
          if ($t.rABS) {
            $t.wb = XLSX.read(btoa(this.fixdata(data)), {  // 手动转化
              type: 'base64'
            })
          } else {
            $t.wb = XLSX.read(data, {
              type: 'binary'
            })
          }
          let json = XLSX.utils.sheet_to_json($t.wb.Sheets[$t.wb.SheetNames[0]])
          console.log(typeof json)
          $t.dealFile($t.analyzeData(json)) // analyzeData: 解析导入数据
        }
        if (this.rABS) {
          reader.readAsArrayBuffer(f)
        } else {
          reader.readAsBinaryString(f)
        }
      },
      downloadExl: function (json, downName, type) {  // 导出到excel
        let keyMap = [] // 获取键
        for (let k in json[0]) {
          keyMap.push(k)
        }
        console.info('keyMap', keyMap, json)
        let tmpdata = [] // 用来保存转换好的json
        json.map((v, i) => keyMap.map((k, j) => Object.assign({}, {
          v: v[k],
          position: (j > 25 ? this.getCharCol(j) : String.fromCharCode(65 + j)) + (i + 1)
        }))).reduce((prev, next) => prev.concat(next)).forEach(function (v) {
          tmpdata[v.position] = {
            v: v.v
          }
        })
        let outputPos = Object.keys(tmpdata)  // 设置区域,比如表格从A1到D10
        let tmpWB = {
          SheetNames: ['mySheet'], // 保存的表标题
          Sheets: {
            'mySheet': Object.assign({},
              tmpdata, // 内容
              {
                '!ref': outputPos[0] + ':' + outputPos[outputPos.length - 1] // 设置填充区域
              })
          }
        }
        let tmpDown = new Blob([this.s2ab(XLSX.write(tmpWB,
          {bookType: (type === undefined ? 'xlsx' : type), bookSST: false, type: 'binary'} // 这里的数据是用来定义导出的格式类型
        ))], {
          type: ''
        })  // 创建二进制对象写入转换好的字节流
        var href = URL.createObjectURL(tmpDown)  // 创建对象超链接
        this.outFile.download = downName + '.xlsx'  // 下载名称
        this.outFile.href = href  // 绑定a标签
        this.outFile.click()  // 模拟点击实现下载
        setTimeout(function () {  // 延时释放
          URL.revokeObjectURL(tmpDown) // 用URL.revokeObjectURL()来释放这个object URL
        }, 100)
      },
      analyzeData: function (data) {  // 此处可以解析导入数据
        return data
      },
      dealFile: function (data) {   // 处理导入的数据
        console.log(data)
        this.imFile.value = ''
        this.fullscreenLoading = false
        if (data.length <= 0) {
          this.errorDialog = true
          this.errorMsg = '请导入正确信息'
        } else {
          this.excelData = [...this.excelData, ...data]
        }
      },
      s2ab: function (s) { // 字符串转字符流
        var buf = new ArrayBuffer(s.length)
        var view = new Uint8Array(buf)
        for (var i = 0; i !== s.length; ++i) {
          view[i] = s.charCodeAt(i) & 0xFF
        }
        return buf
      },
      getCharCol: function (n) { // 将指定的自然数转换为26进制表示。映射关系：[0-25] -> [A-Z]。
        let s = ''
        let m = 0
        while (n > 0) {
          m = n % 26 + 1
          s = String.fromCharCode(m + 64) + s
          n = (n - m) / 26
        }
        return s
      },
      fixdata: function (data) {  // 文件流转BinaryString
        var o = ''
        var l = 0
        var w = 10240
        for (; l < data.byteLength / w; ++l) {
          o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)))
        }
        o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)))
        return o
      },
      exportSQLFile: _.debounce(function() {
        this.log('test')
        if (!this.selType) {
          this.$message({
            showClose: true,
            message: '请先选择sql类别!',
            type: 'warning',
            duration: 1000
          });
          return
        }
        // INSERT INTO i18ndictionary (`uuid`, `tenant`, `lang`, `type`, `keyId`, `keyValue`, `note`, `queue`) VALUES (UUID(), '123', 'ZH', '系统', '系统/公告管理/详情/供应商', '供应商', NULL, NULL);
        // INSERT INTO i18ndictionary (`uuid`, `tenant`, `lang`, `type`, `keyId`, `keyValue`, `note`, `queue`) VALUES (UUID(), '123', 'EN', '系统', '系统/公告管理/详情/供应商', 'Supplier', NULL, NULL);
        let type = '系统'
        this.excelData.forEach(item => {
          let normalz = 'INSERT IGNORE INTO i18ndictionary (`uuid`, `tenant`, `lang`, `type`, `keyId`, `keyValue`, `note`, `queue`)' +  ` VALUES (UUID(), '123', 'ZH', '${this.selType}', '${item.keyid}', '${item.ZH}', NULL, NULL);`
          let normale = 'INSERT IGNORE INTO i18ndictionary (`uuid`, `tenant`, `lang`, `type`, `keyId`, `keyValue`, `note`, `queue`)' +  ` VALUES (UUID(), '123', 'EN', '${this.selType}', '${item.keyid}', '${item.EN}', NULL, NULL);`
          this.sqlArr.push(normalz)
          this.sqlArr.push(normale)
        })
      }, 200)
    }
  }
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
  .index {
    text-align: center;
    color: #2c3e50;
    margin-top: 60px;
  }
  .el-table {
    margin-bottom: 20px;
  }
  .el-table th>.cell {
    text-align: left;
  }
  .button {
    margin-bottom: 20px;
  }
  .sqlCss{
    text-align: left;
  }
</style>
