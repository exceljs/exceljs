# ExcelJS

[![Build Status](https://travis-ci.org/exceljs/exceljs.svg?branch=master)](https://travis-ci.org/exceljs/exceljs)
[![Code Quality: Javascript](https://img.shields.io/lgtm/grade/javascript/g/exceljs/exceljs.svg?logo=lgtm&logoWidth=18)](https://lgtm.com/projects/g/exceljs/exceljs/context:javascript)
[![Total Alerts](https://img.shields.io/lgtm/alerts/g/exceljs/exceljs.svg?logo=lgtm&logoWidth=18)](https://lgtm.com/projects/g/exceljs/exceljs/alerts)

读取，操作和编写电子表格数据和样式到XLSX和JSON。

从Excel电子表格文件逆向工程设计的项目。

# 安装

```shell
npm install exceljs
```

# 新功能！

<ul>
  <li>
    合并的 <a href="https://github.com/exceljs/exceljs/pull/834">添加中文文档 #834</a>.
    非常感谢 <a href="https://github.com/loverto">flydragon</a> 为此贡献。
  </li>
</ul>

# 贡献

非常欢迎贡献！它帮助我了解所需的功能或哪些错误导致最大的痛点。

我只有一个请求;如果您提交了错误修正的拉取请求，请在spec文件夹中添加捕获问题的单元测试或集成测试。
即使只有失败测试的公关也很好 - 我可以分析测试的内容并从中修复代码。

需要明确的是，添加到此库中的所有贡献都将包含在库的MIT许可证中。

# 待完善列表

<ul>
  <li>ESLint  - 慢慢开启（合理的）规则，我希望这些规则应该有助于提高贡献。</li>
  <li>条件格式。</li>
  <li>还有更多的打印设置要添加;固定行/列等</li>
  <li>XLSX流式读。</li>
  <li>用标题解析CSV</li>
</ul>

# 目录

<ul>
  <li><a href="#importing">导入</a></li>
  <li>
    <a href="#interface">接口</a>
    <ul>
      <li><a href="#create-a-workbook">创建工作簿</a></li>
      <li><a href="#set-workbook-properties">设置工作簿属性</a></li>
      <li><a href="#workbook-views">工作簿视图</a></li>
      <li><a href="#add-a-worksheet">添加工作表</a></li>
      <li><a href="#remove-a-worksheet">删除工作表</a></li>
      <li><a href="#access-worksheets">访问工作表</a></li>
      <li><a href="#worksheet-state">工作表状态</a></li>
      <li><a href="#worksheet-properties">工作表属性</a></li>
      <li><a href="#page-setup">页面设置</a></li>
      <li><a href="#header-footer">页眉和页脚</a></li>
      <li>
        <a href="#worksheet-views">工作表视图</a>
        <ul>
          <li><a href="#frozen-views">冻结视图</a></li>
          <li><a href="#split-views">拆分视图</a></li>
        </ul>
      </li>
      <li><a href="#auto-filters">自动过滤器</a></li>
      <li><a href="#columns">列</a></li>
      <li><a href="#rows">行</a></li>
      <li><a href="#handling-individual-cells">处理单个单元格</a></li>
      <li><a href="#merged-cells">合并单元格</a></li>
      <li><a href="#defined-names">定义名称</a></li>
      <li><a href="#data-validations">数据验证</a></li>
      <li><a href="#styles">样式</a>
        <ul>
          <li><a href="#number-formats">数字格式</a></li>
          <li><a href="#fonts">字体</a></li>
          <li><a href="#alignment">对齐方式</a></li>
          <li><a href="#borders">边框</a></li>
          <li><a href="#fills">填充</a></li>
          <li><a href="#rich-text">富文本</a></li>
        </ul>
      </li>
      <li><a href="#outline-levels">大纲级别</a></li>
      <li><a href="#images">图片</a></li>
      <li><a href="#file-io">文件 I/O</a>
        <ul>
          <li><a href="#xlsx">XLSX</a>
            <ul>
              <li><a href="#reading-xlsx">读 XLSX</a></li>
              <li><a href="#writing-xlsx">写 XLSX</a></li>
            </ul>
          </li>
          <li><a href="#csv">CSV</a>
            <ul>
              <li><a href="#reading-csv">读 CSV</a></li>
              <li><a href="#writing-csv">写 CSV</a></li>
            </ul>
          </li>
          <li><a href="#streaming-io">流 I/O</a>
            <ul>
              <li><a href="#streaming-xlsx">流 XLSX</a></li>
            </ul>
          </li>
        </ul>
      </li>
    </ul>
  </li>
  <li><a href="#browser">浏览器</a></li>
  <li>
    <a href="#value-types">值类型</a>
    <ul>
      <li><a href="#null-value">空值</a></li>
      <li><a href="#merge-cell">合并单元格</a></li>
      <li><a href="#number-value">数值</a></li>
      <li><a href="#string-value">字符串值</a></li>
      <li><a href="#date-value">日期值</a></li>
      <li><a href="#hyperlink-value">超链接值</a></li>
      <li>
        <a href="#formula-value">公式值</a>
        <ul>
          <li><a href="#shared-formula">共享公式</a></li>
          <li><a href="#formula-type">公式类型</a></li>
        </ul>
      </li>
      <li><a href="#rich-text-value">富文本值</a></li>
      <li><a href="#boolean-value">布尔值</a></li>
      <li><a href="#error-value">错误值</a></li>
    </ul>
  </li>
  <li><a href="#config">配置</a></li>
  <li><a href="#known-issues">已知的问题</a></li>
  <li><a href="#release-history">发布历史</a></li>
</ul>

## <a id="importing">导入</a>

默认导出是带有 Promise polyfill 转换的 ES5 版本 - 因为这会提供最高的兼容性。

```javascript
var Excel = require('exceljs');
import Excel from 'exceljs';
```

但是，如果您在现代node.js版本（>=8）上使用此库，或者在使用bundler的前端（或者只关注常绿浏览器）上使用此库，我们建议使用这些导入：

```javascript
const Excel = require('exceljs/modern.nodejs');
import Excel from 'exceljs/modern.browser';
```

# <a id="interface">接口</a>

## <a id="create-a-workbook">创建工作簿</a>

```javascript
var workbook = new Excel.Workbook();
```

## <a id="set-workbook-properties">设置工作簿属性</a>

```javascript
workbook.creator = 'Me';
workbook.lastModifiedBy = 'Her';
workbook.created = new Date(1985, 8, 30);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2016, 9, 27);
```

```javascript
// 将工作簿日期设置为1904日期系统
workbook.properties.date1904 = true;
```

## <a id="workbook-views">工作簿视图</a>

“工作簿”视图控制Excel在查看工作簿时打开多少个单独的窗口。

```javascript
workbook.views = [
  {
    x: 0, y: 0, width: 10000, height: 20000,
    firstSheet: 0, activeTab: 1, visibility: 'visible'
  }
]
```

## <a id="add-a-worksheet">添加工作表</a>

```javascript
var sheet = workbook.addWorksheet('My Sheet');
```

用addWorksheet函数的第二个参数设置工作表的选项。

例如

```javascript
// 创建一个红色标签颜色的工作表
var sheet = workbook.addWorksheet('My Sheet', {properties:{tabColor:{argb:'FFC0000'}}});

// 创建一个隐藏网格线的工作表
var sheet = workbook.addWorksheet('My Sheet', {views: [{showGridLines: false}]});

// 创建一个第一行和列冻结的工作表
var sheet = workbook.addWorksheet('My Sheet', {views:[{xSplit: 1, ySplit:1}]});
```

## <a id="remove-a-worksheet">删除工作表</a>

使用工作表`id`从工作簿中删除工作表。

例如

```javascript
// 创建工作表
var sheet = workbook.addWorksheet('My Sheet');

// 使用工作表ID删除工作表
workbook.removeWorksheet(sheet.id)
```

## <a id="access-worksheets">访问工作表</a>
```javascript
// 迭代所有sheet
// 注意：workbook.worksheets.forEach仍然可以工作，但这个方式更好
workbook.eachSheet(function(worksheet, sheetId) {
  // ...
});

// 按名称获取表格
var worksheet = workbook.getWorksheet('My Sheet');

// 按ID获取表格
var worksheet = workbook.getWorksheet(1);
```

## <a id="worksheet-state">工作表状态</a>

```javascript
// 使工作表可见
worksheet.state = 'visible';

// 使工作表隐藏
worksheet.state = 'hidden';

// 使工作表隐藏在“隐藏/取消隐藏”对话框中
worksheet.state = 'veryHidden';
```

## <a id="worksheet-properties">工作表属性</a>

工作表支持属性桶，以允许控制工作表的某些功能。

```javascript
// 创建具有属性的新工作表
var worksheet = workbook.addWorksheet('sheet', {properties:{tabColor:{argb:'FF00FF00'}}});

// 创建一个具有属性的可写的新工作表
var worksheetWriter = workbookWriter.addSheet('sheet', {properties:{outlineLevelCol:1}});

// 之后调整属性（不受工作表 - 编写者支持）
worksheet.properties.outlineLevelCol = 2;
worksheet.properties.defaultRowHeight = 15;
```

**支持的属性**

| 名称             | 默认值    | 描述 |
| ---------------- | ---------- | ----------- |
| tabColor         | undefined  | 标签的颜色 |
| outlineLevelCol  | 0          | 工作表列大纲级别 |
| outlineLevelRow  | 0          | 工作表行大纲级别 |
| defaultRowHeight | 15         | 默认行高 |
| defaultColWidth  | (可选) | 默认列宽 |
| dyDescent        | 55         | TBD |

### 工作表指标

一些新的指标已添加到工作表...

| 名称              | 描述 |
| ----------------- | ----------- |
| rowCount          | 文档的总行大小。等于有值的最后一行的行号。 |
| actualRowCount    | 具有值的行数的计数。如果中间文档行为空，则它不会包含在计数中。 |
| columnCount       | 文档的总列大小。等于所有行的最大单元格计数 |
| actualColumnCount | 具有值的列数的计数。 |


## <a id="page-setup">页面设置</a>

所有可能影响工作表打印的属性都保存在工作表的pageSetup对象中。

```javascript
// 使用A4设置的页面设置设置创建新工作表 - 横向
var worksheet =  workbook.addWorksheet('sheet', {
  pageSetup:{paperSize: 9, orientation:'landscape'}
});

// 使用适合页面的pageSetup设置创建一个新的工作表编写器
var worksheetWriter = workbookWriter.addSheet('sheet', {
  pageSetup:{fitToPage: true, fitToHeight: 5, fitToWidth: 7}
});

// 之后调整pageSetup设置
worksheet.pageSetup.margins = {
  left: 0.7, right: 0.7,
  top: 0.75, bottom: 0.75,
  header: 0.3, footer: 0.3
};

// 设置工作表的打印区域
worksheet.pageSetup.printArea = 'A1:G20';

// 在每个打印页面上重复特定行
worksheet.pageSetup.printTitlesRow = '1:3';

// 在每个打印页面上重复特定列
worksheet.pageSetup.printTitlesColumn = 'A:C';
```

**支持的页面设置设置**

| 属性名                  | 默认值       | 描述 |
| --------------------- | ------------- | ----------- |
| margins               |               | 页面边框上的空格。单位是英寸。 |
| orientation           | 'portrait'    | 页面的方向 - 即更高（纵向）或更宽（横向） |
| horizontalDpi         | 4294967295    | 每英寸水平点。默认值为-1 |
| verticalDpi           | 4294967295    | 每英寸垂直点。默认值为-1 |
| fitToPage             |               | 是否使用fitToWidth和fitToHeight或缩放设置。默认值基于pageSetup对象中是否存在这些设置 - 如果两者都存在，则缩放获胜（即默认值为false） |
| pageOrder             | 'downThenOver'| 打印页面的顺序 - 其中之一 ['downThenOver', 'overThenDown'] |
| blackAndWhite         | false         | 打印没有颜色 |
| draft                 | false         | 打印质量较差（和墨水） |
| cellComments          | 'None'        | 在哪里发表评论 - 其中之一 ['atEnd', 'asDisplayed', 'None'] |
| errors                | 'displayed'   | 在哪里显示错误 - 其中之一 ['dash', 'blank', 'NA', 'displayed'] |
| scale                 | 100           | 增加或减小打印尺寸的百分比值。当fitToPage为false时激活 |
| fitToWidth            | 1             | 页面应打印多少页宽。当fitToPage为true时激活  |
| fitToHeight           | 1             | 页面应打印多少页。当fitToPage为true时激活  |
| paperSize             |               | 使用什么纸张尺寸（见下文） |
| showRowColHeaders     | false         | 是否显示行号和列字母 |
| showGridLines         | false         | 是否显示网格线 |
| firstPageNumber       |               | 第一页使用哪个号码 |
| horizontalCentered    | false         | 是否水平居中工作表数据 |
| verticalCentered      | false         | 是否垂直居中表格数据 |

**示例纸张尺寸**

| 名称                          | 值     |
| ----------------------------- | --------- |
| 信                        | undefined |
| Legal                         |  5        |
| Executive                     |  7        |
| A4                            |  9        |
| A5                            |  11       |
| B5 (JIS)                      |  13       |
| 信封 #10                       |  20       |
| 信封 DL                        |  27       |
| 信封 C5                        |  28       |
| 信封 B5                        |  34       |
| 信封君主                       |  37       |
| 日本双明信片旋转 |  82       |
| 16K 197x273 mm                |  119      |

## <a id="header-footer">页眉和页脚</a>

这里将介绍如何添加页眉和页脚，添加的内容主要是文本，比如时间，简介，文件信息等，并且可以设置文本的风格。此外，也可以针对首页，奇偶页设置不同的文本。

警告：不支持添加图片

```javascript
// 代码中出现的&开头字符对应变量，相关信息可查阅下文的变量表
// 设置页脚(默认居中),结果：“第 2 页，共 16 页”
worksheet.headerFooter.oddFooter = "第 &P 页，共 &N 页";

// 设置页脚(默认居中)加粗,结果：“第 2 页，共 16 页”
worksheet.headerFooter.oddFooter = "&B第 &P 页，共 &N 页";

// 设置左边页脚为18px大小并倾斜,结果：“第 2 页，共 16 页”
worksheet.headerFooter.oddFooter = "&L&18&I第 &P 页，共 &N 页";

// 设置中间页眉为灰色微软雅黑,结果：“52 exceljs”
worksheet.headerFooter.oddHeader = "&C&KCCCCCC&\"微软雅黑\"52 exceljs";

// 设置页脚的左中右文本，结果：页脚左“exceljs” 页脚中“demo.xlsx” 页脚右“第 2 页”
worksheet.headerFooter.oddFooter = "&Lexceljs&C&F&R第 &P 页";

// 为首页设置独特的内容
worksheet.headerFooter.differentFirst = true;
worksheet.headerFooter.firstHeader = "Hello Exceljs";
worksheet.headerFooter.firstFooter = "Hello World"
```

**属性表**

| 名称              | 默认值   | 描述 |
| ----------------- | --------- | ----------- |
|differentFirst|false|开启或关闭首页使用独特的文本内容|
|differentOddEven|false|开启或关闭奇数页和偶数页使用不同的文本内容|
|oddHeader|null|奇数页的页眉内容，如果 differentOddEven = false ，那么该项作为页面默认的页眉内容|
|oddFooter|null|奇数页的页脚内容，如果 differentOddEven = false ，那么该项作为页面默认的页脚内容|
|evenHeader|null|偶数页的页眉内容，differentOddEven = true 后有效|
|evenFooter|null|偶数页的页脚内容，differentOddEven = true 后有效|
|firstHeader|null|首页的页眉内容，differentFirst = true 后有效|
|firstFooter|null|首页的页脚内容，differentFirst = true 后有效|

**变量表**

| 名称                | 描述 |
| -----------------  | ----------- |
|&L|设置位置为左边|
|&C|设置位置为中间|
|&R|设置位置为右边|
|&P|当前页数|
|&N|总页数|
|&D|当前日期|
|&T|当前时间|
|&G|图片|
|&A|工作表名|
|&F|文件名|
|&B|文本加粗|
|&I|文本倾斜|
|&U|文本下划线|
|&"字体名称"|字体名称，比如：&“微软雅黑”|
|&数字|字体大小，比如12px大的文本：&12 |
|&KHEXCode|字体颜色，比如灰色：&KCCCCCC|

## <a id="worksheet-views">工作表视图</a>

工作表现在支持一个视图列表，用于控制Excel如何显示工作表：

* 冻结 - 顶部和左侧的多个行和列被冻结到位。只有左下部分会滚动
* split - 视图分为4个部分，每个部分半独立可滚动。

每个视图还支持各种属性：

| 名称              | 默认值   | 描述 |
| ----------------- | --------- | ----------- |
| state             | 'normal'  | 控制视图状态，其中任意一个 normal, frozen or split |
| rightToLeft       | false     | 将工作表视图的方向设置为 right-to-left |
| activeCell        | undefined | 当前选择的单元格 |
| showRuler         | true      | 在页面布局中显示或隐藏标尺 |
| showRowColHeaders | true      | 显示或隐藏行和列标题（例如顶部的A1，B1和左侧的1,2,3 |
| showGridLines     | true      | 显示或隐藏网格线（显示未定义边框的单元格） |
| zoomScale         | 100       | 用于视图的缩放百分比 |
| zoomScaleNormal   | 100       | 正常缩放视图 |
| style             | undefined | 演示文稿样式 -  pageBreakPreview或pageLayout之一。注意pageLayout与冻结视图不兼容 |

### <a id="frozen-views">冰冻视图</a>

冻结视图支持以下扩展属性：

| 名称              | 默认值   | 描述 |
| ----------------- | --------- | ----------- |
| xSplit            | 0         | 要冻结的列数。要仅冻结行，请将其设置为0或未定义 |
| ySplit            | 0         | 要冻结的行数。要仅冻结列，请将其设置为0或未定义 |
| topLeftCell       | special   | 哪个单元格将位于右下方窗格的左上角。注意：不能是冻结单元格。默认为首先解冻单元格 |

```javascript
worksheet.views = [
  {state: 'frozen', xSplit: 2, ySplit: 3, topLeftCell: 'G10', activeCell: 'A1'}
];
```

### <a id="split-views">拆分视图</a>

拆分视图支持以下扩展属性：

| 名称              | 默认值   | 描述 |
| ----------------- | --------- | ----------- |
| xSplit            | 0         | 从左边多少个点设置拆分点。要垂直拆分，请将其设置为0或未定义 |
| ySplit            | 0         | 从顶部多少个点设置拆分点。要水平拆分，请将其设置为0或未定义  |
| topLeftCell       | undefined | 哪个单元格将位于右下方窗格的左上角。 |
| activePane        | undefined | 哪个窗格将处于活动状态 -  topLeft，topRight，bottomLeft和bottomRight之一 |

```javascript
worksheet.views = [
  {state: 'split', xSplit: 2000, ySplit: 3000, topLeftCell: 'G10', activeCell: 'A1'}
];
```

## <a id="auto-filters">自动过滤器</a>

可以将自动过滤器应用于工作表。

```javascript
worksheet.autoFilter = 'A1:C1';
```

虽然范围字符串是autoFilter的标准形式，但工作表也将支持
以下值：

```javascript
// 设置从A1到C1的自动过滤器
worksheet.autoFilter = {
  from: 'A1',
  to: 'C1',
}

// 设置从第3行的第1列中的单元格到第5行的第12列中的单元格的自动过滤器
worksheet.autoFilter = {
  from: {
    row: 3,
    column: 1
  },
  to: {
    row: 5,
    column: 12
  }
}

// 将自动过滤器从D3设置为第7行和第5列中的单元格
worksheet.autoFilter = {
  from: 'D3',
  to: {
    row: 7,
    column: 5
  }
}
```

## <a id="columns">列</a>

```javascript
// 添加列标题并定义列键和宽度
// 注意：这些列结构只是工作簿构建方便，
// 除了列宽之外，它们不会完全持久化。
worksheet.columns = [
  { header: 'Id', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 32 },
  { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 }
];

// 按键，字母和从1开始的列号访问各列
var idCol = worksheet.getColumn('id');
var nameCol = worksheet.getColumn('B');
var dobCol = worksheet.getColumn(3);

// 设置列属性

// 注意：将覆盖单元格值C1
dobCol.header = 'Date of Birth';

// 注意：这将覆盖单元格值C1:C2
dobCol.header = ['Date of Birth', 'A.K.A. D.O.B.'];

// 从这一点开始，此列将被'dob'索引而不是'DOB'
dobCol.key = 'dob';

dobCol.width = 15;

// 如果您愿意，请隐藏该列
dobCol.hidden = true;

// 设置列的大纲级别
worksheet.getColumn(4).outlineLevel = 0;
worksheet.getColumn(5).outlineLevel = 1;

// columns支持只读字段，以指示基于outlineLevel的折叠状态
expect(worksheet.getColumn(4).collapsed).to.equal(false);
expect(worksheet.getColumn(5).collapsed).to.equal(true);

// 迭代此列中的所有当前单元格
dobCol.eachCell(function(cell, rowNumber) {
  // ...
});

// 迭代此列中的所有当前单元格，包括空单元格
dobCol.eachCell({ includeEmpty: true }, function(cell, rowNumber) {
  // ...
});

// 添加一列新值
worksheet.getColumn(6).values = [1,2,3,4,5];

// 添加稀疏的值列
worksheet.getColumn(7).values = [,,2,3,,5,,7,,,,11];

// 剪切一列或多列（右侧列向左移动）
// 如果已定义列属性，则会相应地剪切或移动它们
//已知问题：如果拼接导致任何合并的单元格移动，则结果可能无法预测
worksheet.spliceColumns(3,2);

// 删除一列并再插入两列。
// 注意：第4列及以上将向右移动1列。
// 另外：如果工作表的行数多于column插入中的值，则行仍将被移位，就像存在的值一样
var newCol3Values = [1,2,3,4,5];
var newCol4Values = ['one', 'two', 'three', 'four', 'five'];
worksheet.spliceColumns(3, 1, newCol3Values, newCol4Values);

```

## <a id="rows">行</a>

```javascript
// 使用列键在最后一行之后按键值添加几行
worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970,1,1)});
worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965,1,7)});

// 按连续数组添加一行（分配给A，B和C列）
worksheet.addRow([3, 'Sam', new Date()]);

// 通过稀疏数组添加一行（分配给A，E和I列）
var rowValues = [];
rowValues[1] = 4;
rowValues[5] = 'Kyle';
rowValues[9] = new Date();
worksheet.addRow(rowValues);

// 添加行数组
var rows = [
  [5,'Bob',new Date()], // row by array
  {id:6, name: 'Barbara', dob: new Date()}
];
worksheet.addRows(rows);

// 获取行对象。如果它尚不存在，将返回一个新的空的
var row = worksheet.getRow(5);

// 获取工作表中的最后一个可编辑行（如果没有，则为undefined）
var row = worksheet.lastRow;

// 设置特定的行高
row.height = 42.5;

// 使行隐藏
row.hidden = true;

// 设置行的大纲级别
worksheet.getRow(4).outlineLevel = 0;
worksheet.getRow(5).outlineLevel = 1;

// rows支持readonly字段以指示基于outlineLevel的折叠状态
expect(worksheet.getRow(4).collapsed).to.equal(false);
expect(worksheet.getRow(5).collapsed).to.equal(true);


row.getCell(1).value = 5; // A5的值设为5
row.getCell('name').value = 'Zeb'; // B5的值设置为'Zeb' - 假设第2列仍然按名称键入
row.getCell('C').value = new Date(); // C5的值设定为当前日期

// 获取一行作为稀疏数组
// 注意：接口变更：worksheet.getRow(4)==> worksheet.getRow(4).values
row = worksheet.getRow(4).values;
expect(row[5]).toEqual('Kyle');

// 通过连续数组分配行值（其中数组元素0具有值）
row.values = [1,2,3];
expect(row.getCell(1).value).toEqual(1);
expect(row.getCell(2).value).toEqual(2);
expect(row.getCell(3).value).toEqual(3);

// 通过稀疏数组分配行值（其中数组元素0未定义）
var values = []
values[5] = 7;
values[10] = 'Hello, World!';
row.values = values;
expect(row.getCell(1).value).toBeNull();
expect(row.getCell(5).value).toEqual(7);
expect(row.getCell(10).value).toEqual('Hello, World!');

// 使用列键按对象分配行值
row.values = {
  id: 13,
  name: 'Thing 1',
  dob: new Date()
};

// 在行下方插入分页符
row.addPageBreak();

// 迭代工作表中具有值的所有行
worksheet.eachRow(function(row, rowNumber) {
  console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
});

// 迭代工作表中的所有行（包括空行）
worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
  console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
});

// 迭代连续的所有非空单元格
row.eachCell(function(cell, colNumber) {
  console.log('Cell ' + colNumber + ' = ' + cell.value);
});

// 迭代连续的所有单元格（包括空单元格）
row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
  console.log('Cell ' + colNumber + ' = ' + cell.value);
});

// 剪切一行或多行（下面的行向上移动）
// 已知问题：如果拼接导致任何合并的单元格移动，则结果可能无法预测
worksheet.spliceRows(4,3);

// 删除一行并再插入两行。
//注意：第4行和第4行将向下移动1行。
var newRow3Values = [1,2,3,4,5];
var newRow4Values = ['one', 'two', 'three', 'four', 'five'];
worksheet.spliceRows(3, 1, newRow3Values, newRow4Values);

// 切一个或多个单元格（右边的单元格向左移动）
// 注意：此操作不会影响其他行
row.splice(3,2);

// 移除一个单元格并再插入两个单元格（切割单元格右侧的单元格将向右移动）
row.splice(4,1,'new value 1', 'new value 2');

// 提交已完成的行进行流式处理
row.commit();

// 行指标
var rowSize = row.cellCount;
var numValues = row.actualCellCount;
```

## <a id="handling-individual-cells">处理单个单元格</a>

```javascript
var cell = worksheet.getCell('C3');

// 修改/添加单个单元格
cell.value = new Date(1968, 5, 1);

// 查询单元格的类型
expect(cell.type).toEqual(Excel.ValueType.Date);

// 使用单元格的字符串值
myInput.value = cell.text;

// 使用html-safe字符串进行渲染...
var html = '<div>' + cell.html + '</div>';

```

## <a id="merged-cells">合并单元格</a>

```javascript
// 合并一系列单元格
worksheet.mergeCells('A4:B5');

// ... 合并的单元格是链接的
worksheet.getCell('B5').value = 'Hello, World!';
expect(worksheet.getCell('B5').value).toBe(worksheet.getCell('A4').value);
expect(worksheet.getCell('B5').master).toBe(worksheet.getCell('A4'));

// ... 合并的单元格共享相同的样式对象
expect(worksheet.getCell('B5').style).toBe(worksheet.getCell('A4').style);
worksheet.getCell('B5').style.font = myFonts.arial;
expect(worksheet.getCell('A4').style.font).toBe(myFonts.arial);

// 取消单元格破坏风格链接
worksheet.unMergeCells('A4');
expect(worksheet.getCell('B5').style).not.toBe(worksheet.getCell('A4').style);
expect(worksheet.getCell('B5').style.font).not.toBe(myFonts.arial);

// 由左上角，右下角合并
worksheet.mergeCells('G10', 'H11');
worksheet.mergeCells(10,11,12,13); // 上，左，下，右
```

## <a id="defined-names">定义的名称</a>

单个单元格（或多组单元格）可以为其分配名称。
 名称可用于公式和数据验证（可能更多）。

```javascript
// 为单元格分配（或获取）名称（将覆盖单元格具有的任何其他名称）
worksheet.getCell('A1').name = 'PI';
expect(worksheet.getCell('A1').name).to.equal('PI');

// 为单元格分配（或获取）一组名称（单元格可以有多个名称）
worksheet.getCell('A1').names = ['thing1', 'thing2'];
expect(worksheet.getCell('A1').names).to.have.members(['thing1', 'thing2']);

// 从单元格中删除名称
worksheet.getCell('A1').removeName('thing1');
expect(worksheet.getCell('A1').names).to.have.members(['thing2']);
```

## <a id="data-validations">数据验证</a>

单元格可以定义哪些值有效，并向用户提供提示以帮助指导它们。

验证类型可以是以下之一：

| Type       | 描述 |
| ---------- | ----------- |
| list       | 定义一组离散的有效值。 Excel将在下拉列表中提供这些内容以便于输入 |
| whole      | 值必须是整数 |
| decimal    | 该值必须是十进制数 |
| textLength | 值可以是文本，但长度是受控的 |
| custom     | 自定义公式控制有效值 |

对于list或custom以外的类型，以下运算符会影响验证：

| Operator              | 描述 |
| --------------------  | ----------- |
| between               | 值必须位于公式结果之间 |
| notBetween            | 值不能介于公式结果之间 |
| equal                 | 值必须等于公式结果 |
| notEqual              | 值必须不等于公式结果 |
| greaterThan           | 值必须大于公式结果 |
| lessThan              | 值必须小于公式结果 |
| greaterThanOrEqual    | 值必须大于或等于公式结果 |
| lessThanOrEqual       | 值必须小于或等于公式结果 |

```javascript
//指定有效值列表（一，二，三，四）。
// Excel将提供包含这些值的下拉列表。
worksheet.getCell('A1').dataValidation = {
  type: 'list',
  allowBlank: true,
  formulae: ['"One,Two,Three,Four"']
};

// 指定范围中的有效值列表。
// Excel将提供包含这些值的下拉列表。
worksheet.getCell('A1').dataValidation = {
  type: 'list',
  allowBlank: true,
  formulae: ['$D$5:$F$5']
};

// 指定单元格必须是不是5的整数。
// 如果用户输入错误，请向用户显示相应的错误消息
worksheet.getCell('A1').dataValidation = {
  type: 'whole',
  operator: 'notEqual',
  showErrorMessage: true,
  formulae: [5],
  errorStyle: 'error',
  errorTitle: 'Five',
  error: 'The value must not be Five'
};

// Specify Cell must be a decimal number between 1.5 and 7.
// Add 'tooltip' to help guid the user
worksheet.getCell('A1').dataValidation = {
  type: 'decimal',
  operator: 'between',
  allowBlank: true,
  showInputMessage: true,
  formulae: [1.5, 7],
  promptTitle: 'Decimal',
  prompt: 'The value must between 1.5 and 7'
};

// Specify Cell must be have a text length less than 15
worksheet.getCell('A1').dataValidation = {
  type: 'textLength',
  operator: 'lessThan',
  showErrorMessage: true,
  allowBlank: true,
  formulae: [15]
};

// Specify Cell must be have be a date before 1st Jan 2016
worksheet.getCell('A1').dataValidation = {
  type: 'date',
  operator: 'lessThan',
  showErrorMessage: true,
  allowBlank: true,
  formulae: [new Date(2016,0,1)]
};
```

## <a id="cell-comments">单元格评论</a>

将旧样式评论添加到单元格

```javascript
// 纯文本评论
worksheet.getCell('A1').note = 'Hello, ExcelJS!';

// 彩色格式的评论
ws.getCell('B1').note = {
  texts: [
    {'font': {'size': 12, 'color': {'theme': 0}, 'name': 'Calibri', 'family': 2, 'scheme': 'minor'}, 'text': 'This is '},
    {'font': {'italic': true, 'size': 12, 'color': {'theme': 0}, 'name': 'Calibri', 'scheme': 'minor'}, 'text': 'a'},
    {'font': {'size': 12, 'color': {'theme': 1}, 'name': 'Calibri', 'family': 2, 'scheme': 'minor'}, 'text': ' '},
    {'font': {'size': 12, 'color': {'argb': 'FFFF6600'}, 'name': 'Calibri', 'scheme': 'minor'}, 'text': 'colorful'},
    {'font': {'size': 12, 'color': {'theme': 1}, 'name': 'Calibri', 'family': 2, 'scheme': 'minor'}, 'text': ' text '},
    {'font': {'size': 12, 'color': {'argb': 'FFCCFFCC'}, 'name': 'Calibri', 'scheme': 'minor'}, 'text': 'with'},
    {'font': {'size': 12, 'color': {'theme': 1}, 'name': 'Calibri', 'family': 2, 'scheme': 'minor'}, 'text': ' in-cell '},
    {'font': {'bold': true, 'size': 12, 'color': {'theme': 1}, 'name': 'Calibri', 'family': 2, 'scheme': 'minor'}, 'text': 'format'},
  ],
};
```

## <a id="styles">样式</a>

单元格，行和列各自支持一组丰富的样式和格式，这些样式和格式会影响单元格的显示方式。

通过分配以下属性来设置样式：

* <a href="#number-formats">数字格式</a>
* <a href="#fonts">字体</a>
* <a href="#alignment">对齐方式</a>
* <a href="#borders">边框</a>
* <a href="#fills">填充</a>

```javascript
// assign a style to a cell
ws.getCell('A1').numFmt = '0.00%';

// Apply styles to worksheet columns
ws.columns = [
  { header: 'Id', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 32, style: { font: { name: 'Arial Black' } } },
  { header: 'D.O.B.', key: 'DOB', width: 10, style: { numFmt: 'dd/mm/yyyy' } }
];

// Set Column 3 to Currency Format
ws.getColumn(3).numFmt = '"£"#,##0.00;[Red]\-"£"#,##0.00';

// Set Row 2 to Comic Sans.
ws.getRow(2).font = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true };
```

将样式应用于行或列时，它将应用于该行或列中的所有当前现有单元格。
 此外，创建的任何新单元格都将从其所属的行和列继承其初始样式。

如果单元格的行和列都定义了特定样式（例如字体），则单元格将使用行样式而不是列样式。
 但是，如果行和列定义了不同的样式（例如column.numFmt和row.font），则单元格将继承行中的字体和列中的numFmt。

警告：所有上述属性（numFmt除外，它是一个字符串）都是JS对象结构。
 如果将同一样式对象分配给多个电子表格实体，则每个实体将共享相同的样式对象。
 如果在序列化电子表格之前稍后修改了样式对象，则也将修改引用该样式对象的所有实体。
 此行为旨在通过减少创建的JS对象的数量来确定性能的优先级。
 如果希望样式对象是独立的，则需要在分配它们之前克隆它们。
 此外，默认情况下，如果从文件（或流）读取文档（如果电子表格实体共享相似的样式），则它们也将引用相同的样式对象。

### <a id="number-formats">数字格式</a>

```javascript
// 显示值为 '1 3/5'
ws.getCell('A1').value = 1.6;
ws.getCell('A1').numFmt = '# ?/?';

// 显示值为 '1.60%'
ws.getCell('B1').value = 0.016;
ws.getCell('B1').numFmt = '0.00%';
```

### <a id="fonts">字体</a>

```javascript

// 对于想要那些想要的平面设计师
ws.getCell('A1').font = {
  name: 'Comic Sans MS',
  family: 4,
  size: 16,
  underline: true,
  bold: true
};

// 为毕业生平面设计师...
ws.getCell('A2').font = {
  name: 'Arial Black',
  color: { argb: 'FF00FF00' },
  family: 2,
  size: 14,
  italic: true
};

// 用于垂直对齐
ws.getCell('A3').font = {
  vertAlign: 'superscript'
};

// 注意：单元格将存储对指定的字体对象的引用。
// 如果之后更改了字体对象，则单元格字体也将更改...
var font = { name: 'Arial', size: 12 };
ws.getCell('A3').font = font;
font.size = 20; // Cell A3现在的字体大小为20！

// 从文件或流中读取工作簿后，共享相似字体的单元格可能会引用相同的字体对象
```

| Font Property | 描述       | Example Value(s) |
| ------------- | ----------------- | ---------------- |
| name          | 字体名称。 | 'Arial', 'Calibri', etc. |
| family        | Font family for fallback. An integer value. | 1 - Serif, 2 - Sans Serif, 3 - Mono, Others - unknown |
| scheme        | Font scheme. | 'minor', 'major', 'none' |
| charset       | Font charset. An integer value. | 1, 2, etc. |
| size          | Font size. An integer value. | 9, 10, 12, 16, etc. |
| color         | Colour description, an object containing an ARGB value. | { argb: 'FFFF0000'} |
| bold          | Font **weight** | true, false |
| italic        | Font *slope* | true, false |
| underline     | Font <u>underline</u> style | true, false, 'none', 'single', 'double', 'singleAccounting', 'doubleAccounting' |
| strike        | Font <strike>strikethrough</strike> | true, false |
| outline       | Font outline | true, false |
| vertAlign     | Vertical align | 'superscript', 'subscript'

### <a id="alignment">对齐方式</a>

```javascript
// set cell alignment to top-left, middle-center, bottom-right
ws.getCell('A1').alignment = { vertical: 'top', horizontal: 'left' };
ws.getCell('B1').alignment = { vertical: 'middle', horizontal: 'center' };
ws.getCell('C1').alignment = { vertical: 'bottom', horizontal: 'right' };

// set cell to wrap-text
ws.getCell('D1').alignment = { wrapText: true };

// set cell indent to 1
ws.getCell('E1').alignment = { indent: 1 };

// set cell text rotation to 30deg upwards, 45deg downwards and vertical text
ws.getCell('F1').alignment = { textRotation: 30 };
ws.getCell('G1').alignment = { textRotation: -45 };
ws.getCell('H1').alignment = { textRotation: 'vertical' };
```

**有效的对齐属性值**

| horizontal       | vertical    | wrapText | indent  | readingOrder | textRotation |
| ---------------- | ----------- | -------- | ------- | ------------ | ------------ |
| left             | top         | true     | integer | rtl          | 0 to 90      |
| center           | middle      | false    |         | ltr          | -1 to -90    |
| right            | bottom      |          |         |              | vertical     |
| fill             | distributed |          |         |              |              |
| justify          | justify     |          |         |              |              |
| centerContinuous |             |          |         |              |              |
| distributed      |             |          |         |              |              |


### <a id="borders">边框</a>

```javascript
// 在A1周围设置单个细边框
ws.getCell('A1').border = {
  top: {style:'thin'},
  left: {style:'thin'},
  bottom: {style:'thin'},
  right: {style:'thin'}
};

// 在A3附近设置双细绿边框
ws.getCell('A3').border = {
  top: {style:'double', color: {argb:'FF00FF00'}},
  left: {style:'double', color: {argb:'FF00FF00'}},
  bottom: {style:'double', color: {argb:'FF00FF00'}},
  right: {style:'double', color: {argb:'FF00FF00'}}
};

// 在A5设置厚红十字
ws.getCell('A5').border = {
  diagonal: {up: true, down: true, style:'thick', color: {argb:'FFFF0000'}}
};
```

**有效的边框样式**

* thin
* dotted
* dashDot
* hair
* dashDotDot
* slantDashDot
* mediumDashed
* mediumDashDotDot
* mediumDashDot
* medium
* double
* thick

### <a id="fills">填充</a>

```javascript
// fill A1 with red darkVertical stripes
ws.getCell('A1').fill = {
  type: 'pattern',
  pattern:'darkVertical',
  fgColor:{argb:'FFFF0000'}
};

// fill A2 with yellow dark trellis and blue behind
ws.getCell('A2').fill = {
  type: 'pattern',
  pattern:'darkTrellis',
  fgColor:{argb:'FFFFFF00'},
  bgColor:{argb:'FF0000FF'}
};

// fill A3 with blue-white-blue gradient from left to right
ws.getCell('A3').fill = {
  type: 'gradient',
  gradient: 'angle',
  degree: 0,
  stops: [
    {position:0, color:{argb:'FF0000FF'}},
    {position:0.5, color:{argb:'FFFFFFFF'}},
    {position:1, color:{argb:'FF0000FF'}}
  ]
};


// fill A4 with red-green gradient from center
ws.getCell('A4').fill = {
  type: 'gradient',
  gradient: 'path',
  center:{left:0.5,top:0.5},
  stops: [
    {position:0, color:{argb:'FFFF0000'}},
    {position:1, color:{argb:'FF00FF00'}}
  ]
};
```

#### <a id="pattern-fills">图案填充</a>

| Property | Required | 描述 |
| -------- | -------- | ----------- |
| type     | Y        | Value: 'pattern'<br/>Specifies this fill uses patterns |
| pattern  | Y        | Specifies type of pattern (see <a href="#valid-pattern-types">Valid Pattern Types</a> below) |
| fgColor  | N        | Specifies the pattern foreground color. Default is black. |
| bgColor  | N        | Specifies the pattern background color. Default is white. |

**Valid Pattern Types**

* none
* solid
* darkGray
* mediumGray
* lightGray
* gray125
* gray0625
* darkHorizontal
* darkVertical
* darkDown
* darkUp
* darkGrid
* darkTrellis
* lightHorizontal
* lightVertical
* lightDown
* lightUp
* lightGrid
* lightTrellis

#### <a id="gradient-fills">渐变填充</a>

| Property | Required | 描述 |
| -------- | -------- | ----------- |
| type     | Y        | Value: 'gradient'<br/>Specifies this fill uses gradients |
| gradient | Y        | Specifies gradient type. One of ['angle', 'path'] |
| degree   | angle    | For 'angle' gradient, specifies the direction of the gradient. 0 is from the left to the right. Values from 1 - 359 rotates the direction clockwise |
| center   | path     | For 'path' gradient. Specifies the relative coordinates for the start of the path. 'left' and 'top' values range from 0 to 1 |
| stops    | Y        | Specifies the gradient colour sequence. Is an array of objects containing position and color starting with position 0 and ending with position 1. Intermediary positions may be used to specify other colours on the path. |

**注意事项**

使用上面的界面可以创建使用XLSX编辑器程序无法实现的渐变填充效果。
例如，Excel仅支持0度，45度，90度和135度的角度渐变。
类似地，停止序列也可以由具有位置[0,1]或[0,0.5,1]作为唯一选项的UI限制。
注意此填充以确保目标XLSX查看器支持它。

### <a id="rich-text">富文本</a>

单个单元格现在支持富文本或单元格格式。
 富文本值可以控制文本值中任意数量子字符串的字体属性。
 有关支持哪些字体属性的详细信息的完整列表，请参见<a href="font">字体</a>。

```javascript

ws.getCell('A1').value = {
  'richText': [
    {'font': {'size': 12,'color': {'theme': 0},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'This is '},
    {'font': {'italic': true,'size': 12,'color': {'theme': 0},'name': 'Calibri','scheme': 'minor'},'text': 'a'},
    {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' '},
    {'font': {'size': 12,'color': {'argb': 'FFFF6600'},'name': 'Calibri','scheme': 'minor'},'text': 'colorful'},
    {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' text '},
    {'font': {'size': 12,'color': {'argb': 'FFCCFFCC'},'name': 'Calibri','scheme': 'minor'},'text': 'with'},
    {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' in-cell '},
    {'font': {'bold': true,'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'format'}
  ]
};

expect(ws.getCell('A1').text).to.equal('This is a colorful text with in-cell format');
expect(ws.getCell('A1').type).to.equal(Excel.ValueType.RichText);

```

## <a id="outline-levels">大纲级别</a>

Excel支持概述;其中可以展开或折叠行或列，具体取决于用户希望查看的详细程度。

可以在列设置中定义大纲级别：
```javascript
worksheet.columns = [
  { header: 'Id', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 32 },
  { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 }
];
```

或直接在行或列上
```javascript
worksheet.getColumn(3).outlineLevel = 1;
worksheet.getRow(3).outlineLevel = 1;
```

可以在工作表上设置工作表大纲级别
```javascript
// 设置列大纲级别
worksheet.properties.outlineLevelCol = 1;

// set row outline level
worksheet.properties.outlineLevelRow = 1;
```

注意：调整行或列上的大纲级别或工作表上的大纲级别将产生副作用，即还修改受属性更改影响的所有行或列的折叠属性。例如。：
```javascript
worksheet.properties.outlineLevelCol = 1;

worksheet.getColumn(3).outlineLevel = 1;
expect(worksheet.getColumn(3).collapsed).to.be.true;

worksheet.properties.outlineLevelCol = 2;
expect(worksheet.getColumn(3).collapsed).to.be.false;
```

可以在工作表上设置轮廓属性

```javascript
worksheet.properties.outlineProperties = {
  summaryBelow: false,
  summaryRight: false,
};
```

## <a id="images">图片</a>

将图像添加到工作表需要两个步骤。
首先，通过addImage（）函数将图像添加到工作簿，该函数也将返回imageId值。
然后，使用imageId，可以将图像作为平铺背景或覆盖单元格范围添加到工作表中。

注意：从此版本开始，不支持调整或转换图像。

### <a id="add-image-to-workbook">将图像添加到工作簿</a>

Workbook.addImage函数支持按文件名或缓冲区添加图像。
请注意，在这两种情况下，都必须指定扩展名。
有效的扩展值包括'jpeg'，'png'，'gif'。

```javascript
// add image to workbook by filename
var imageId1 = workbook.addImage({
  filename: 'path/to/image.jpg',
  extension: 'jpeg',
});

// add image to workbook by buffer
var imageId2 = workbook.addImage({
  buffer: fs.readFileSync('path/to.image.png'),
  extension: 'png',
});

// add image to workbook by base64
var myBase64Image = "data:image/png;base64,iVBORw0KG...";
var imageId2 = workbook.addImage({
  base64: myBase64Image,
  extension: 'png',
});
```

### <a id="add-image-background-to-worksheet">将图像背景添加到工作表</a>

使用Workbook.addImage中的图像ID，可以使用addBackgroundImage函数设置工作表的背景

```javascript
// set background
worksheet.addBackgroundImage(imageId1);
```

### <a id="add-image-over-a-range">在范围内添加图像</a>

使用Workbook.addImage中的图像ID，可以在工作表中嵌入图像以覆盖范围。
从该范围计算的坐标将覆盖从第一个单元格的左上角到第二个单元格的右下角。

```javascript
// insert an image over B2:D6
worksheet.addImage(imageId2, 'B2:D6');
```

使用结构而不是范围字符串，可以部分覆盖单元格。

请注意，用于此的坐标系统为零，因此A1的左上角将为{col：0，row：0}。
可以通过使用浮点数来指定单元的分数，例如， A1的中点是{col：0.5，row：0.5}。

```javascript
// insert an image over part of B2:D6
worksheet.addImage(imageId2, {
  tl: { col: 1.5, row: 1.5 },
  br: { col: 3.5, row: 5.5 }
});
```

单元格范围也可以具有“editAs”属性，它将控制图像如何锚定到单元格
它可以具有以下值之一：

| Value     | 描述 |
| --------- | ----------- |
| undefined | 这是默认值。它指定将使用单元格移动和调整图像大小 |
| oneCell   | 图像将与单元格一起移动但不是大小 |
| absolute  | 图像不会随单元格移动或调整大小 |

```javascript
ws.addImage(imageId, {
  tl: { col: 0.1125, row: 0.4 },
  br: { col: 2.101046875, row: 3.4 },
  editAs: 'oneCell'
});
```

### <a id="add-image-to-a-cell">将图像添加到单元格</a>

You can add an image to a cell and then define its width and height in pixels at 96dpi.

```javascript
worksheet.addImage(imageId2, {
  tl: { col: 0, row: 0 },
  ext: { width: 500, height: 200 }
});
```

### <a id="add-image-with-hyperlinks">添加带超链接的图片</a>

You can add an image with hyperlinks to a cell.

```javascript
worksheet.addImage(imageId2, {
  tl: { col: 0, row: 0 },
  ext: { width: 500, height: 200 },
  hyperlinks: {
    hyperlink: 'http://www.somewhere.com',
    tooltip: 'http://www.somewhere.com'
  }
});
```

## <a id="file-io">文件 I/O</a>

### <a id="xlsx">XLSX</a>

#### <a id="reading-xlsx">读 XLSX</a>

```javascript
// read from a file
var workbook = new Excel.Workbook();
workbook.xlsx.readFile(filename)
  .then(function() {
    // use workbook
  });

// pipe from stream
var workbook = new Excel.Workbook();
stream.pipe(workbook.xlsx.createInputStream());

// load from buffer
var workbook = new Excel.Workbook();
workbook.xlsx.load(data)
  .then(function() {
    // use workbook
  });
```

#### <a id="writing-xlsx">写 XLSX</a>

```javascript
// write to a file
var workbook = createAndFillWorkbook();
workbook.xlsx.writeFile(filename)
  .then(function() {
    // done
  });

// write to a stream
workbook.xlsx.write(stream)
  .then(function() {
    // done
  });

// write to a new buffer
workbook.xlsx.writeBuffer()
  .then(function(buffer) {
    // done
  });
```

### CSV <a id="csv">CSV</a>

#### <a id="reading-csv">读 CSV</a>

```javascript
// read from a file
var workbook = new Excel.Workbook();
workbook.csv.readFile(filename)
  .then(worksheet => {
    // use workbook or worksheet
  });

// read from a stream
var workbook = new Excel.Workbook();
workbook.csv.read(stream)
  .then(worksheet => {
    // use workbook or worksheet
  });

// pipe from stream
var workbook = new Excel.Workbook();
stream.pipe(workbook.csv.createInputStream());

// read from a file with European Dates
var workbook = new Excel.Workbook();
var options = {
  dateFormats: ['DD/MM/YYYY']
};
workbook.csv.readFile(filename, options)
  .then(worksheet => {
    // use workbook or worksheet
  });

// read from a file with custom value parsing
var workbook = new Excel.Workbook();
var options = {
  map(value, index) {
    switch(index) {
      case 0:
        // column 1 is string
        return value;
      case 1:
        // column 2 is a date
        return new Date(value);
      case 2:
        // column 3 is JSON of a formula value
        return JSON.parse(value);
      default:
        // the rest are numbers
        return parseFloat(value);
    }
  }
};
workbook.csv.readFile(filename, options)
  .then(function(worksheet) {
    // use workbook or worksheet
  });
```

CSV解析器使用[fast-csv](https://www.npmjs.com/package/fast-csv)来读取CSV文件。
 传递给上述读取函数的选项也传递给fast-csv以解析csv数据。
 有关详细信息，请参阅fast-csv README.md。

使用npm模块[moment](https://www.npmjs.com/package/moment)解析日期。
 如果未提供dateFormats，则使用以下内容：

* moment.ISO_8601
*'MM-DD-YYYY'
*'YYYY-MM-DD'

#### <a id="writing-csv">写 CSV</a>

```javascript

// write to a file
var workbook = createAndFillWorkbook();
workbook.csv.writeFile(filename)
  .then(() => {
    // done
  });

// write to a stream
// Be careful that you need to provide sheetName or
// sheetId for correct import to csv.
workbook.csv.write(stream, { sheetName: 'Page name' })
  .then(() => {
    // done
  });

// write to a file with European Date-Times
var workbook = new Excel.Workbook();
var options = {
  dateFormat: 'DD/MM/YYYY HH:mm:ss',
  dateUTC: true, // use utc when rendering dates
};
workbook.csv.writeFile(filename, options)
  .then(() => {
    // done
  });


// write to a file with custom value formatting
var workbook = new Excel.Workbook();
var options = {
  map(value, index) {
    switch(index) {
      case 0:
        // column 1 is string
        return value;
      case 1:
        // column 2 is a date
        return moment(value).format('YYYY-MM-DD');
      case 2:
        // column 3 is a formula, write just the result
        return value.result;
      default:
        // the rest are numbers
        return value;
    }
  }
};
workbook.csv.writeFile(filename, options)
  .then(() => {
    // done
  });

// write to a new buffer
workbook.csv.writeBuffer()
  .then(function(buffer) {
    // done
  });
```

CSV解析器使用[fast-csv]（https://www.npmjs.com/package/fast-csv）编写CSV文件。
 传递给上述写入函数的选项也传递给fast-csv以写入csv数据。
 有关详细信息，请参阅fast-csv README.md。

使用npm模块格式化日期[时刻]（https://www.npmjs.com/package/moment）。
 如果未提供dateFormat，则使用moment.ISO_8601。
 编写CSV时，您可以将布尔值dateUTC设置为true，以使ExcelJS自动解析日期
 使用`moment.utc（）`转换时区。

### <a id="streaming-io">流 I/O</a>

上面记录的文件I / O要求在写入文件之前在内存中构建整个工作簿。
 虽然方便，但由于所需的内存量，它可能会限制文档的大小。

流编写器（或读取器）在生成时处理工作簿或工作表数据，
 将其转换为文件格式。通常情况下，这在内存上的效率要高得多
 内存占用，甚至中间内存占用比文档版本更紧凑，
 特别是当您认为行和单元格对象在提交后处理时。

流工作簿和工作表的界面几乎与文档版本相同，但有一些细微的实际差异：

*将工作表添加到工作簿后，无法将其删除。
*一旦提交了一行，它将不再可访问，因为它将从工作表中删除。
*不支持unMergeCells（）。

请注意，可以在不提交任何行的情况下构建整个工作簿。
 提交工作簿时，将自动提交所有添加的工作表（包括所有未提交的行）。
 但是，在这种情况下，文档版本的收益很少。
 
#### <a id="streaming-xlsx">流 XLSX</a>

##### Streaming XLSX Writer

流式XLSX编写器在ExcelJS.stream.xlsx命名空间中可用。

构造函数使用带有以下字段的可选选项对象：

| Field            | 描述 |
| ---------------- | ----------- |
| stream           | Specifies a writable stream to write the XLSX workbook to. |
| filename         | If stream not specified, this field specifies the path to a file to write the XLSX workbook to. |
| useSharedStrings | Specifies whether to use shared strings in the workbook. Default is false |
| useStyles        | Specifies whether to add style information to the workbook. Styles can add some performance overhead. Default is false |

如果选项中既未指定流也未指定文件名，则工作簿编写器将创建StreamBuf对象
 这将把XLSX工作簿的内容存储在内存中。
 这个StreamBuf对象可以通过属性workbook.stream访问，也可以用于
 通过stream.read（）直接访问字节或将内容传递给另一个流。

```javascript
// construct a streaming XLSX workbook writer with styles and shared strings
var options = {
  filename: './streamed-workbook.xlsx',
  useStyles: true,
  useSharedStrings: true
};
var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
```

通常，流式XLSX编写器的接口与文档工作簿（和工作表）相同
 如上所述，实际上行，单元格和样式对象是相同的。

然而，有一些差异......

**施工**

如上所示，WorkbookWriter通常需要在构造函数中指定输出流或文件。

**提交数据**

当工作表行准备就绪时，应该提交它以便可以释放行对象和内容。
 通常，这将在添加每一行时完成...

```javascript
worksheet.addRow({
   id: i,
   name: theName,
   etc: someOtherDetail
}).commit();
```

WorksheetWriter在添加行时不提交行的原因是允许跨行合并单元格：
```javascript
worksheet.mergeCells('A1:B2');
worksheet.getCell('A1').value = 'I am merged';
worksheet.getCell('C1').value = 'I am not';
worksheet.getCell('C2').value = 'Neither am I';
worksheet.getRow(2).commit(); // now rows 1 and two are committed.
```

每个工作表完成后，还必须提交：

```javascript
// Finished adding data. Commit the worksheet
worksheet.commit();
```

要完成XLSX文档，必须提交工作簿。如果未提交工作簿中的任何工作表，
 它们将作为工作簿提交的一部分自动提交。

```javascript
// Finished the workbook.
workbook.commit()
  .then(function() {
    // the stream has been written
  });
```

# <a id="browser">浏览器</a>

该库的一部分已经过隔离测试，可在浏览器环境中使用。

由于工作簿阅读器和工作簿编写器的流媒体特性，因此未包含这些内容。
只能使用基于文档的工作簿（有关详细信息，请参阅<a href="#create-a-workbook">创建Workbook </a>）。

例如，在浏览器中使用ExcelJS的代码，请查看github中的<a href="https://github.com/exceljs/exceljs/tree/master/spec/browser"> spec / browser </a>文件夹回购。

## 预先捆绑

以下文件已预先捆绑并包含在dist文件夹中。

* exceljs.js
* exceljs.min.js

# <a id="value-types">值类型</a>

支持以下值类型。

## <a id="null-value">空值</a>

Enum: Excel.ValueType.Null

空值表示缺少值，并且在写入文件时通常不会存储（合并单元格除外）。
  它可用于从单元格中删除值。

例如

```javascript
worksheet.getCell('A1').value = null;
```

## <a id="merge-cell">合并单元格</a>

Enum: Excel.ValueType.Merge

合并单元格的值与另一个“主”单元格绑定。
  分配给合并单元将导致修改主单元。

## <a id="number-value">数值</a>

Enum: Excel.ValueType.Number

A numeric value.

E.g.

```javascript
worksheet.getCell('A1').value = 5;
worksheet.getCell('A2').value = 3.14159;
```

## <a id="string-value">字符串值</a>

Enum: Excel.ValueType.String

A simple text string.

E.g.

```javascript
worksheet.getCell('A1').value = 'Hello, World!';
```

## <a id="date-value">日期值</a>

Enum: Excel.ValueType.Date

A date value, represented by the JavaScript Date type.

E.g.

```javascript
worksheet.getCell('A1').value = new Date(2017, 2, 15);
```

## <a id="hyperlink-value">超链接值</a>

Enum: Excel.ValueType.Hyperlink

A URL with both text and link value.

E.g.
```javascript
// link to web
worksheet.getCell('A1').value = {
  text: 'www.mylink.com',
  hyperlink: 'http://www.mylink.com',
  tooltip: 'www.mylink.com'
};

// internal link
worksheet.getCell('A1').value = { text: 'Sheet2', hyperlink: '#\'Sheet2\'!A1' };
```

## <a id="formula-value">公式值</a>

Enum: Excel.ValueType.Formula

An Excel formula for calculating values on the fly.
  Note that while the cell type will be Formula, the cell may have an effectiveType value that will
  be derived from the result value.

Note that ExcelJS cannot process the formula to generate a result, it must be supplied.

E.g.

```javascript
worksheet.getCell('A3').value = { formula: 'A1+A2', result: 7 };
```

Cells also support convenience getters to access the formula and result:

```javascript
worksheet.getCell('A3').formula === 'A1+A2';
worksheet.getCell('A3').result === 7;
```

### <a id="shared-formula">共享公式</a>

Shared formulae enhance the compression of the xlsx document by increasing the repetition
of text within the worksheet xml.

A shared formula can be assigned to a cell using a new value form:

```javascript
worksheet.getCell('B3').value = { sharedFormula: 'A3', result: 10 };
```

This specifies that the cell B3 is a formula that will be derived from the formula in
A3 and its result is 10.

The formula convenience getter will translate the formula in A3 to what it should be in B3:

```javascript
worksheet.getCell('B3').formula === 'B1+B2';
```

### <a id="formula-type">公式类型</a>

To distinguish between real and translated formula cells, use the formulaType getter:

```javascript
worksheet.getCell('A3').formulaType === Enums.FormulaType.Master;
worksheet.getCell('B3').formulaType === Enums.FormulaType.Shared;
```

Formula type has the following values:

| Name                       |  Value  |
| -------------------------- | ------- |
| Enums.FormulaType.None     |   0     |
| Enums.FormulaType.Master   |   1     |
| Enums.FormulaType.Shared   |   2     |

## <a id="rich-text-value">富文本值</a>

枚举： Excel.ValueType.RichText

富文本格式。

例如
```javascript
worksheet.getCell('A1').value = {
  richText: [
    { text: 'This is '},
    {font: {italic: true}, text: 'italic'},
  ]
};
```

## <a id="boolean-value">布尔值</a>

枚举：Excel.ValueType.Boolean

例如

```javascript
worksheet.getCell('A1').value = true;
worksheet.getCell('A2').value = false;
```

## <a id="error-value">错误值</a>

枚举：Excel.ValueType.Error

例如

```javascript
worksheet.getCell('A1').value = { error: '#N/A' };
worksheet.getCell('A2').value = { error: '#VALUE!' };
```

当前有效的错误文本值为：

| Name                           | Value       |
| ------------------------------ | ----------- |
| Excel.ErrorValue.NotApplicable | #N/A        |
| Excel.ErrorValue.Ref           | #REF!       |
| Excel.ErrorValue.Name          | #NAME?      |
| Excel.ErrorValue.DivZero       | #DIV/0!     |
| Excel.ErrorValue.Null          | #NULL!      |
| Excel.ErrorValue.Value         | #VALUE!     |
| Excel.ErrorValue.Num           | #NUM!       |

# 接口变化

我们尽一切努力建立一个良好的一致界面，不会突破版本，但令人遗憾的是，现在有些事情必须改变以获得更大的利益。

## 0.1.0

### Worksheet.eachRow

Worksheet.eachRow的回调函数中的参数已被交换和更改;它是函数（rowNumber，rowValues），现在它是函数（row，rowNumber），它使它看起来更像是下划线（_.each）函数，并将行对象优先于行号。

### Worksheet.getRow

此函数已从返回稀疏的单元格值数组更改为返回Row对象。这样可以访问行属性，并有助于管理行样式等。

通过Worksheet.getRow（rowNumber）.values仍然可以获得稀疏的单元格值数组;

## 0.1.1

### cell.model

cell.styles重命名为cell.style

## 0.2.44

从Bluebird切换到本机节点Promise的函数返回Promise，如果它们依赖Bluebird的额外功能，它可以破坏调用代码。

为了缓解这种情况，将以下两个更改添加到0.3.0：

* 默认情况下使用功能更全面且仍然与浏览器兼容的promise lib。这个lib支持Bluebird的许多功能，但占用空间要小得多。
* 注入不同Promise实现的选项。有关详细信息，请参阅<a href="#config">配置</a>部分。

# <a id="config">配置</a>

ExcelJS现在支持promise库的依赖注入。
您可以通过在模块中包含以下代码来恢复Bluebird承诺...

```javascript
ExcelJS.config.setValue('promise', require('bluebird'));
```

请注意：我已经专门用蓝鸟测试了ExcelJS（因为直到最近这是它使用的库）。
从我所做的测试来看，它不适用于Q.

# 注意事项

## Dist文件夹

在发布此模块之前，源代码将被转换并以其他方式处理，然后再放入dist/文件夹中。
本自述文件识别两个文件 - 浏览器化的捆绑包和缩小版本。
除了package.json中指定为“main”的文件之外，不保证dist/文件夹的其他任何内容


# <a id="known-issues">已知的问题</a>

## 使用Puppeteer进行测试

此lib中包含的测试套件包括在无头浏览器中执行的小脚本，以验证捆绑的包。在撰写本文时，似乎该测试在Windows Linux子系统中不能很好地发挥作用。

因此，可以通过名为.disable-test-browser的文件来禁用浏览器测试

```bash
sudo apt-get install libfontconfig
```

## Splice vs Merge

如果任何拼接操作影响合并的单元格，则不会正确移动合并组

# <a id="release-history">发布历史</a>

| Version | Changes |
| ------- | ------- |
| 0.0.9   | <ul><li><a href="#number-formats">Number Formats</a></li></ul> |
| 0.1.0   | <ul><li>Bug Fixes<ul><li>"&lt;" and "&gt;" text characters properly rendered in xlsx</li></ul></li><li><a href="#columns">Better Column control</a></li><li><a href="#rows">Better Row control</a></li></ul> |
| 0.1.1   | <ul><li>Bug Fixes<ul><li>More textual data written properly to xml (including text, hyperlinks, formula results and format codes)</li><li>Better date format code recognition</li></ul></li><li><a href="#fonts">Cell Font Style</a></li></ul> |
| 0.1.2   | <ul><li>Fixed potential race condition on zip write</li></ul> |
| 0.1.3   | <ul><li><a href="#alignment">Cell Alignment Style</a></li><li><a href="#rows">Row Height</a></li><li>Some Internal Restructuring</li></ul> |
| 0.1.5   | <ul><li>Bug Fixes<ul><li>Now handles 10 or more worksheets in one workbook</li><li>theme1.xml file properly added and referenced</li></ul></li><li><a href="#borders">Cell Borders</a></li></ul> |
| 0.1.6   | <ul><li>Bug Fixes<ul><li>More compatible theme1.xml included in XLSX file</li></ul></li><li><a href="#fills">Cell Fills</a></li></ul> |
| 0.1.8   | <ul><li>Bug Fixes<ul><li>More compatible theme1.xml included in XLSX file</li><li>Fixed filename case issue</li></ul></li><li><a href="#fills">Cell Fills</a></li></ul> |
| 0.1.9   | <ul><li>Bug Fixes<ul><li>Added docProps files to satisfy Mac Excel users</li><li>Fixed filename case issue</li><li>Fixed worksheet id issue</li></ul></li><li><a href="#set-workbook-properties">Core Workbook Properties</a></li></ul> |
| 0.1.10  | <ul><li>Bug Fixes<ul><li>Handles File Not Found error</li></ul></li><li><a href="#csv">CSV Files</a></li></ul> |
| 0.1.11  | <ul><li>Bug Fixes<ul><li>Fixed Vertical Middle Alignment Issue</li></ul></li><li><a href="#styles">Row and Column Styles</a></li><li><a href="#rows">Worksheet.eachRow supports options</a></li><li><a href="#rows">Row.eachCell supports options</a></li><li><a href="#columns">New function Column.eachCell</a></li></ul> |
| 0.2.0   | <ul><li><a href="#streaming-xlxs-writer">Streaming XLSX Writer</a><ul><li>At long last ExcelJS can support writing massive XLSX files in a scalable memory efficient manner. Performance has been optimised and even smaller spreadsheets can be faster to write than the document writer. Options have been added to control the use of shared strings and styles as these can both have a considerable effect on performance</li></ul></li><li><a href="#rows">Worksheet.lastRow</a><ul><li>Access the last editable row in a worksheet.</li></ul></li><li><a href="#rows">Row.commit()</a><ul><li>For streaming writers, this method commits the row (and any previous rows) to the stream. Committed rows will no longer be editable (and are typically deleted from the worksheet object). For Document type workbooks, this method has no effect.</li></ul></li></ul> |
| 0.2.2   | <ul><li><a href="https://pbs.twimg.com/profile_images/2933552754/fc8c70829ee964c5542ae16453503d37.jpeg">One Billion Cells</a><ul><li>Achievement Unlocked: A simple test using ExcelJS has created a spreadsheet with 1,000,000,000 cells. Made using random data with 100,000,000 rows of 10 cells per row. I cannot validate the file yet as Excel will not open it and I have yet to implement the streaming reader but I have every confidence that it is good since 1,000,000 rows loads ok.</li></ul></li></ul> |
| 0.2.3   | <ul><li>Bug Fixes<ul><li><a href="https://github.com/exceljs/exceljs/issues/18">Merge Cell Styles</a><ul><li>Merged cells now persist (and parse) their styles.</li></ul></li></ul></li><li><a href="#streaming-xlxs-writer">Streaming XLSX Writer</a><ul><li>At long last ExcelJS can support writing massive XLSX files in a scalable memory efficient manner. Performance has been optimised and even smaller spreadsheets can be faster to write than the document writer. Options have been added to control the use of shared strings and styles as these can both have a considerable effect on performance</li></ul></li><li><a href="#rows">Worksheet.lastRow</a><ul><li>Access the last editable row in a worksheet.</li></ul></li><li><a href="#rows">Row.commit()</a><ul><li>For streaming writers, this method commits the row (and any previous rows) to the stream. Committed rows will no longer be editable (and are typically deleted from the worksheet object). For Document type workbooks, this method has no effect.</li></ul></li></ul> |
| 0.2.4   | <ul><li>Bug Fixes<ul><li><a href="https://github.com/exceljs/exceljs/issues/27">Worksheets with Ampersand Names</a><ul><li>Worksheet names are now xml-encoded and should work with all xml compatible characters</li></ul></li></ul></li><li><a href="#rows">Row.hidden</a> & <a href="#columns">Column.hidden</a><ul><li>Rows and Columns now support the hidden attribute.</li></ul></li><li><a href="#worksheet">Worksheet.addRows</a><ul><li>New function to add an array of rows (either array or object form) to the end of a worksheet.</li></ul></li></ul> |
| 0.2.6   | <ul><li>Bug Fixes<ul><li><a href="https://github.com/exceljs/exceljs/issues/87">invalid signature: 0x80014</a>: Thanks to <a href="https://github.com/hasanlussa">hasanlussa</a> for the PR</li></ul></li><li><a href="#defined-names">Defined Names</a><ul><li>Cells can now have assigned names which may then be used in formulas.</li></ul></li><li>Converted Bluebird.defer() to new Bluebird(function(resolve, reject){}). Thanks to user <a href="https://github.com/Nishchit14">Nishchit</a> for the Pull Request</li></ul> |
| 0.2.7   | <ul><li><a href="#data-validations">Data Validations</a><ul><li>Cells can now define validations that controls the valid values the cell can have</li></ul></li></ul> |
| 0.2.8   | <ul><li><a href="rich-text">Rich Text Value</a><ul><li>Cells now support <b><i>in-cell</i></b> formatting - Thanks to <a href="https://github.com/pvadam">Peter ADAM</a></li></ul></li><li>Fixed typo in README - Thanks to <a href="https://github.com/MRdNk">MRdNk</a></li><li>Fixing emit in worksheet-reader - Thanks to <a href="https://github.com/alangunning">Alan Gunning</a></li><li>Clearer Docs - Thanks to <a href="https://github.com/miensol">miensol</a></li></ul> |
| 0.2.9   | <ul><li>Fixed "read property 'richText' of undefined error. Thanks to  <a href="https://github.com/james075">james075</a></li></ul> |
| 0.2.10  | <ul><li>Refactoring Complete. All unit and integration tests pass.</li></ul> |
| 0.2.11  | <ul><li><a href="#outline-level">Outline Levels</a>. Thanks to <a href="https://github.com/cricri">cricri</a> for the contribution.</li><li><a href="#worksheet-properties">Worksheet Properties</a></li><li>Further refactoring of worksheet writer.</li></ul> |
| 0.2.12  | <ul><li><a href="#worksheet-views">Sheet Views</a>. Thanks to <a href="https://github.com/cricri">cricri</a> again for the contribution.</li></ul> |
| 0.2.13  | <ul><li>Fix for <a href="https://github.com/exceljs/exceljs/issues">exceljs might be vulnerable for regular expression denial of service</a>. Kudos to <a href="https://github.com/yonjah">yonjah</a> and <a href="https://www.youtube.com/watch?v=wCfE-9bhY2Y">Josh Emerson</a> for the resolution.</li><li>Fix for <a href="https://github.com/exceljs/exceljs/issues/162">Multiple Sheets opens in 'Group' mode in Excel</a>. My bad - overzealous sheet view code.</li><li>Also fix for empty sheet generating invalid xlsx.</li></ul> |
| 0.2.14  | <ul><li>Fix for <a href="https://github.com/exceljs/exceljs/issues">exceljs might be vulnerable for regular expression denial of service</a>. Kudos to <a href="https://github.com/yonjah">yonjah</a> and <a href="https://www.youtube.com/watch?v=wCfE-9bhY2Y">Josh Emerson</a> for the resolution.</li><li>Fixed <a href="https://github.com/exceljs/exceljs/issues/162">Multiple Sheets opens in 'Group' mode in Excel</a> again. Added <a href="#workbook-views">Workbook views</a>.</li><li>Also fix for empty sheet generating invalid xlsx.</li></ul> |
| 0.2.15  | <ul><li>Added <a href="#page-setup">Page Setup Properties</a>. Thanks to <a href="https://github.com/jackkum">Jackkum</a> for the PR</li></ul> |
| 0.2.16  | <ul><li>New <a href="#page-setup">Page Setup</a> Property: Print Area</li></ul> |
| 0.2.17  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/114">Fix a bug on phonetic characters</a>. This fixes an issue related to reading workbooks with phonetic text in. Note phonetic text is not properly supported yet - just properly ignored. Thanks to <a href="https://github.com/zephyrrider">zephyrrider</a> and <a href="https://github.com/gen6033">gen6033</a> for the contribution.</li></ul> |
| 0.2.18  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/175">Fix regression #150: Stream API fails to write XLSX files</a>. Apologies for the regression! Thanks to <a href="https://github.com/danieleds">danieleds</a> for the fix.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/114">Fix a bug on phonetic characters</a>. This fixes an issue related to reading workbooks with phonetic text in. Note phonetic text is not properly supported yet - just properly ignored. Thanks to <a href="https://github.com/zephyrrider">zephyrrider</a> and <a href="https://github.com/gen6033">gen6033</a> for the contribution.</li></ul> |
| 0.2.19  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/119">Update xlsx.js #119</a>. This should make parsing more resilient to open-office documents. Thanks to <a href="https://github.com/nvitaterna">nvitaterna</a> for the contribution.</li></ul> |
| 0.2.20  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/179">Changes from exceljs/exceljs#127 applied to latest version #179</a>. Fixes parsing of defined name values. Thanks to <a href="https://github.com/agdevbridge">agdevbridge</a> and <a href="https://github.com/priitliivak">priitliivak</a> for the contribution.</li></ul> |
| 0.2.21  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/135">color tabs for worksheet-writer #135</a>. Modified the behaviour to print deprecation warning as tabColor has moved into options.properties. Thanks to <a href="https://github.com/ethanlook">ethanlook</a> for the contribution.</li></ul> |
| 0.2.22  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/136">Throw legible error when failing Value.getType() #136</a>. Thanks to <a href="https://github.com/wulfsolter">wulfsolter</a> for the contribution.</li><li>Honourable mention to contributors whose PRs were fixed before I saw them:<ul><li><a href="https://github.com/haoliangyu">haoliangyu</a></li><li><a href="https://github.com/wulfsolter">wulfsolter</a></li></ul></li></ul> |
| 0.2.23  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/137">Fall back to JSON.stringify() if unknown Cell.Type #137</a> with some modification. If a cell value is assigned to an unrecognisable javascript object, the stored value in xlsx and csv files will  be JSON stringified. Note that if the file is read again, no attempt will be made to parse the stringified JSON text. Thanks to <a href="https://github.com/wulfsolter">wulfsolter</a> for the contribution.</li></ul> |
| 0.2.24  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/166">Protect cell fix #166</a>. This does not mean full support for protected cells merely that the parser is not confused by the extra xml. Thanks to <a href="https://github.com/jayflo">jayflo</a> for the contribution.</li></ul> |
| 0.2.25  | <ul><li>Added functions to delete cells, rows and columns from a worksheet. Modelled after the Array splice method, the functions allow cells, rows and columns to be deleted (and optionally inserted). See <a href="#columns">Columns</a> and <a href="#rows">Rows</a> for details.<br />Note: <a href="#splice-vs-merge">Not compatible with cell merges</a></li></ul> |
| 0.2.26  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/184">Update border-xform.js #184</a>Border edges without style will be parsed and rendered as no-border. Thanks to <a href="https://github.com/skumarnk2">skumarnk2</a> for the contribution.</li></ul> |
| 0.2.27  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/187">Pass views to worksheet-writer #187</a>. Now also passes views to worksheet-writer. Thanks to <a href="https://github.com/Temetz">Temetz</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/189">Do not escape xml characters when using shared strings #189</a>. Fixing bug in shared strings. Thanks to <a href="https://github.com/tkirda">tkirda</a> for the contribution.</li></ul> |
| 0.2.28  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/190">Fix tiny bug [Update hyperlink-map.js] #190</a>Thanks to <a href="https://github.com/lszlkss">lszlkss</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/196">fix typo on sheet view showGridLines option #196</a> "showGridlines" should have been "showGridLines". Thanks to <a href="https://github.com/gadiaz1">gadiaz1</a> for the contribution.</li></ul> |
| 0.2.29  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/199">Fire finish event instead of end event on write stream #199</a> and <a href="https://github.com/exceljs/exceljs/pull/200">Listen for finish event on zip stream instead of middle stream #200</a>. Fixes issues with stream completion events. Thanks to <a href="https://github.com/junajan">junajan</a> for the contribution.</li></ul> |
| 0.2.30  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/201">Fix issue #178 #201</a>. Adds the following properties to workbook:<ul><li>title</li><li>subject</li><li>keywords</li><li>category</li><li>description</li><li>company</li><li>manager</li></ul>Thanks to <a href="https://github.com/stavenko">stavenko</a> for the contribution.</li></ul> |
| 0.2.31  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/203">Fix issue #163: the "spans" attribute of the row element is optional #203</a>. Now xlsx parsing will handle documents without row spans. Thanks to <a href="https://github.com/arturas-vitkauskas">arturas-vitkauskas</a> for the contribution.</li></ul> |
| 0.2.32  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/208">Fix issue 206 #208</a>. Fixes issue reading xlsx files that have been printed. Also adds "lastPrinted" property to Workbook. Thanks to <a href="https://github.com/arturas-vitkauskas">arturas-vitkauskas</a> for the contribution.</li></ul> |
| 0.2.33  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/210">Allow styling of cells with no value. #210</a>. Includes Null type cells with style in the rendering parsing. Thanks to <a href="https://github.com/oferns">oferns</a> for the contribution.</li></ul> |
| 0.2.34  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/212">Fix "Unexpected xml node in parseOpen" bug in LibreOffice documents for attributes dc:language and cp:revision #212</a>. Thanks to <a href="https://github.com/jessica-jordan">jessica-jordan</a> for the contribution.</li></ul> |
| 0.2.35  | <ul><li>Fixed <a href="https://github.com/exceljs/exceljs/issues/74">Getting a column/row count #74</a>. <a href="#worksheet-metrics">Worksheet</a> now has rowCount and columnCount properties (and actual variants), <a href="row">Row</a> has cellCount.</li></ul> |
| 0.2.36  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/217">Stream reader fixes #217</a>. Thanks to <a href="https://github.com/kturney">kturney</a> for the contribution.</li></ul> |
| 0.2.37  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/225">Fix output order of Sheet Properties #225</a>. Thanks to <a href="https://github.com/keeneym">keeneym</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/231">remove empty worksheet[0] from _worksheets #231</a>. Thanks to <a href="https://github.com/pookong">pookong</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/232">do not skip empty string in shared strings so that indexes match #232</a>. Thanks again to <a href="https://github.com/pookong">pookong</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/233">use shared strings for streamed writes #233</a>. Thanks again to <a href="https://github.com/pookong">pookong</a> for the contribution.</li></ul> |
| 0.2.38  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/236">Add a comment for issue #216 #236</a>. Thanks to <a href="https://github.com/jsalwen">jsalwen</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/237">Start on support for 1904 based dates #237</a>. Fixed date handling in documents with the 1904 flag set. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li></ul> |
| 0.2.39  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/245">Stops Bluebird warning about unreturned promise #245</a>. Thanks to <a href="https://github.com/robinbullocks4rb">robinbullocks4rb</a> for the contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/247">Added missing dependency: col-cache.js #247</a>. Thanks to <a href="https://github.com/Manish2005">Manish2005</a> for the contribution. </li> </ul> |
| 0.2.42  | <ul><li>Browser Compatible!<ul><li>Well mostly. I have added a browser sub-folder that contains a browserified bundle and an index.js that can be used to generate another. See <a href="#browser">Browser</a> section for details.</li></ul></li><li>Fixed corrupted theme.xml. Apologies for letting that through.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/253">[BUGFIX] data validation formulae undefined #253</a>. Thanks to <a href="https://github.com/jayflo">jayflo</a> for the contribution.</li></ul> |
| 0.2.43  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/255">added a (maybe partial) solution to issue 99. i wasn't able to create an appropriate test #255</a>. This fixes <a href="https://github.com/exceljs/exceljs/issues/99">Too few data or empty worksheet generate malformed excel file #99</a>. Thanks to <a href="https://github.com/mminuti">mminuti</a> for the contribution.</li></ul> |
| 0.2.44  | <ul><li>Reduced Dependencies.<ul><li>Goodbye lodash, goodbye bluebird. Minified bundle is now just over half what it was in the first version.</li></ul></li></ul> |
| 0.2.45  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/256">Sheets with hyperlinks and data validations are corrupted #256</a>. Thanks to <a href="https://github.com/simon-stoic">simon-stoic</a> for the contribution.</li></ul> |
| 0.2.46  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/259">Exclude character controls from XML output. Fixes #234 #262</a>. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/262">Add support for identifier #259</a>. This fixes <a href="https://github.com/exceljs/exceljs/issues/234">Broken XLSX because of "vertical tab" ascii character in a cell #234</a>. Thanks to <a href="https://github.com/NOtherDev">NOtherDev</a> for the contribution.</li></ul> |
| 0.3.0   | <ul><li>Addressed <a href="https://github.com/exceljs/exceljs/issues/266">Breaking change removing bluebird #266</a>. Appologies for any inconvenience.</li><li>Added Promise library dependency injection. See <a href="#config">Config</a> section for more details.</li></ul> |
| 0.3.1   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/279">Update dependencies #279</a>. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/267">Minor fixes for stream handling #267</a>. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li><li>Added automated tests in phantomjs for the browserified code.</li></ul> |
| 0.4.0   | <ul><li>Fixed issue <a href="https://github.com/exceljs/exceljs/issues/278">Boolean cell with value ="true" is returned as 1 #278</a>. The fix involved adding two new Call Value types:<ul><li><a href="#boolean-value">Boolean Value</a></li><li><a href="#error-value">Error Value</a></li></ul>Note: Minor version has been bumped up to 4 as this release introduces a couple of interface changes:<ul><li>Boolean cells previously will have returned 1 or 0 will now return true or false</li><li>Error cells that previously returned a string value will now return an error structure</li></ul></li><li>Fixed issue <a href="https://github.com/exceljs/exceljs/issues/280">Code correctness - setters don't return a value #280</a>.</li><li>Addressed issue <a href="https://github.com/exceljs/exceljs/issues/288">v0.3.1 breaks meteor build #288</a>.</li></ul> |
| 0.4.1   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/285">Add support for cp:contentStatus #285</a>. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/286">Fix Valid characters in XML (allow \n and \r when saving) #286</a>. Thanks to <a href="https://github.com/Rycochet">Rycochet</a> for the contribution.</li><li>Fixed <a href="https://github.com/exceljs/exceljs/issues/275">hyperlink with query arguments corrupts workbook #275</a>. The hyperlink target is not escaped before serialising in the xml.</li></ul> |
| 0.4.2   | <ul><li><p>Addressed the following issues:<ul><li><a href="https://github.com/exceljs/exceljs/issues/290">White text and borders being changed to black #290</a></li><li><a href="https://github.com/exceljs/exceljs/issues/261">Losing formatting/pivot table from loaded file #261</a></li><li><a href="https://github.com/exceljs/exceljs/issues/272">Solid fill become black #272</a></li></ul>These issues are potentially caused by a bug that caused colours with zero themes, tints or indexes to be rendered and parsed incorrectly.</p><p>Regarding themes: the theme files stored inside the xlsx container hold important information regarding colours, styles etc and if the theme information from a loaded xlsx file is lost, the results can be unpredictable and undesirable. To address this, when an ExcelJS Workbook parses an XLSX file, it will preserve any theme files it finds and include them when writing to a new XLSX. If this behaviour is not desired, the Workbook class exposes a clearThemes() function which will drop the theme content. Note that this behaviour is only implemented in the document based Workbook class, not the streamed Reader and Writer.</p></li></ul> |
| 0.4.3   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/294">Support error references in cell ranges #294</a>. Thanks to <a href="https://github.com/holm">holm</a> for the contribution.</li></ul> |
| 0.4.4   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/297">Issue with copied cells #297</a>. This merge adds support for shared formulas. Thanks to <a href="https://github.com/muscapades">muscapades</a> for the contribution.</li></ul> |
| 0.4.6   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/304">Correct spelling #304</a>. Thanks to <a href="https://github.com/toanalien">toanalien</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/304">Added support for auto filters #306</a>. This adds <a href="#auto-filters">Auto Filters</a> to the Worksheet. Thanks to <a href="https://github.com/C4rmond4i">C4rmond4i</a> for the contribution.</li><li>Restored NodeJS 4.0.0 compatability by removing the destructuring code. My apologies for any inconvenience.</li></ul> |
| 0.4.9   | <ul><li>Switching to transpiled code for distribution. This will ensure compatability with 4.0.0 and above from here on. And it will also allow use of much more expressive JS code in the lib folder!</li><li><a href="#images">Basic Image Support!</a>Images can now be added to worksheets either as a tiled background or stretched over a range. Note: other features like rotation, etc. are not supported yet and will reqeuire further work.</li></ul> |
| 0.4.10  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/319">Add missing Office Rels #319</a>. Thanks goes to <a href="https://github.com/mauriciovillalobos">mauriciovillalobos</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/320">Add printTitlesRow Support #320</a> Thanks goes to <a href="https://github.com/psellers89">psellers89</a> for the contribution.</li></ul> |
| 0.4.11  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/327">Avoid error on anchor with no media #327</a>. Thanks goes to <a href="https://github.com/holm">holm</a> for the contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/332">Assortment of fixes for streaming read #332</a>. Thanks goes to <a href="https://github.com/holm">holm</a> for the contribution.</li></ul> |
| 0.4.12  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/334">Don’t set address if hyperlink r:id is undefined #334</a>. Thanks goes to <a href="https://github.com/holm">holm</a> for the contribution.</li></ul> |
| 0.4.13  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/343">Issue 296 #343</a>. This fixes <a href="https://github.com/exceljs/exceljs/issues/296">Issue with writing newlines #296</a>. Thanks goes to <a href="https://github.com/holly-weisser">holly-weisser</a> for the contribution.</li></ul> |
| 0.4.14  | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/350">Syntax highlighting added ✨ #350</a>. Thanks goes to <a href="https://github.com/rmariuzzo">rmariuzzo</a> for the contribution.</li></ul> |
| 0.5.0   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/356">Fix right to left issues #356</a>. Fixes <a href="https://github.com/exceljs/exceljs/issues/72">Add option to RTL file #72</a> and <a href="https://github.com/exceljs/exceljs/issues/126">Adding an option to set RTL worksheet #126</a>. Big thank you to <a href="https://github.com/alitaheri">alitaheri</a> for this contribution.</li></ul> |
| 0.5.1   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/364">Fix #345 TypeError: Cannot read property 'date1904' of undefined #364</a>. This fixes <a href="https://github.com/exceljs/exceljs/issues/345">TypeError: Cannot read property 'date1904' of undefined #345</a>. Thanks to <a href="https://github.com/Diluka">Diluka</a> for this contribution.</li></ul>
| 0.6.0   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/389">Add rowBreaks feature. #389</a>. Thanks to <a href="https://github.com/brucejo75">brucejo75</a> for this contribution.</li></ul> |
| 0.6.1   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/403">Guard null model fields - fix and tests #403</a>. Thanks to <a href="https://github.com/shdd-cjharries">thecjharries</a> for this contribution. Also thanks to <a href="https://github.com/Rycochet">Ryc O'Chet</a> for help with reviewing.</li></ul> |
| 0.6.2   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/396">Add some comments in readme according csv importing #396</a>. Thanks to <a href="https://github.com/Imperat">Michael Lelyakin</a> for this contribution. Also thanks to <a href="https://github.com/planemar">planemar</a> for help with reviewing. This also closes <a href="https://github.com/exceljs/exceljs/issues/395">csv to stream doesn't work #395</a>.</li></ul> |
| 0.7.0   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/407">Impl &lt;xdr:twoCellAnchor editAs=oneCell&gt; #407</a>. Thanks to <a href="https://github.com/Ockejanssen">Ocke Janssen</a> and <a href="https://github.com/kay-ramme">Kay Ramme</a> for this contribution. This change allows control on how images are anchored to cells.</li></ul> |
| 0.7.1   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/423">Don't break when attempting to import a zip file that's not an Excel file (eg. .numbers) #423</a>. Thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. This change makes exceljs more reslilient when opening non-excel files.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/434">Fixes #419 : Updates readme. #434</a>. Thanks to <a href="https://github.com/getsomecoke">Vishnu Kyatannawar</a> for this contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/436">Don't break when docProps/core.xml contains a &lt;cp:version&gt; tag #436</a>. Thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. This change handles core.xml files with empty version tags.</li></ul>
| 0.8.0   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/442">Add Base64 Image support for the .addImage() method #442</a>. Thanks to <a href="https://github.com/jwmann">James W Mann</a> for this contribution.</li><li>Merged <a href="https://github.com/exceljs/exceljs/pull/453">update moment to 2.19.3 #453</a>. Thanks to <a href="https://github.com/cooltoast">Markan Patel</a> for this contribution.</li></ul> |
| 0.8.1   | <ul><li> Merged <a href="https://github.com/exceljs/exceljs/pull/457">Additional information about font family property #457</a>. Thanks to <a href="https://github.com/kayakyakr">kayakyakr</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/459">Fixes #458 #459</a>. This fixes <a href="https://github.com/exceljs/exceljs/issues/458">Add style to column causes it to be hidden #458</a>. Thanks to <a href="https://github.com/AJamesPhillips">Alexander James Phillips</a> for this contribution. </li> </ul> |
| 0.8.2   | <ul><li>Merged <a href="https://github.com/exceljs/exceljs/pull/466">Don't break when loading an Excel file containing a chartsheet #466</a>. Thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/471">Hotfix/sheet order#257 #471</a>. This fixes <a href="https://github.com/exceljs/exceljs/issues/257">Sheet Order #257</a>. Thanks to <a href="https://github.com/robbi">Robbi</a> for this contribution. </li> </ul> |
| 0.8.3   | <ul> <li> Assimilated <a href="https://github.com/exceljs/exceljs/pull/463">fix #79 outdated dependencies in unzip2</a>. Thanks to <a href="https://github.com/jsamr">Jules Sam. Randolph</a> for starting this fix and a really big thanks to <a href="https://github.com/kachkaev">Alexander Kachkaev</a> for finding the final solution. </li> </ul> |
| 0.8.4   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/479">Round Excel date to nearest millisecond when converting to javascript date #479</a>. Thanks to <a href="https://github.com/bjet007">Benoit Jean</a> for this contribution. </li> </ul> |
| 0.8.5   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/485">Bug fix: wb.worksheets/wb.eachSheet caused getWorksheet(0) to return sheet #485</a>. Thanks to <a href="https://github.com/mah110020">mah110020</a> for this contribution. </li> </ul> |
| 0.9.0   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/489">Feature/issue 424 #489</a>. This fixes <a href="https://github.com/exceljs/exceljs/issues/424">No way to control summaryBelow or summaryRight #424</a>. Many thanks to <a href="https://github.com/sarahdmsi">Sarah</a> for this contribution. </li> </ul>  |
| 0.9.1   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/490">add type definition #490</a>. This adds type definitions to ExcelJS! Many thanks to <a href="https://github.com/taoqf">taoqf</a> for this contribution. </li> </ul> |
| 1.0.0   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/494">Add Node 8 and Node 9 to continuous integration testing #494</a>. Many thanks to <a href="https://github.com/cooltoast">Markan Patel</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/508">Small README fix #508</a>. Many thanks to <a href="https://github.com/lbguilherme">Guilherme Bernal</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/501">Add support for inlineStr, including rich text #501</a>. Many thanks to <a href="https://github.com/linguamatics-pdenes">linguamatics-pdenes</a> and <a href="https://github.com/robscotts4rb">Rob Scott</a> for their efforts towards this contribution. Since this change is technically a breaking change (the rendered XML for inline strings will change) I'm making this a major release! </li> </ul> |
| 1.0.1   | <ul> <li> Fixed <a href="https://github.com/exceljs/exceljs/issues/520">spliceColumns problem when the number of columns are important #520</a>. </li> </ul> |
| 1.0.2   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/524">Loosen exceljs's dependency requirements for moment #524</a>. Many thanks to <a href="https://github.com/nicoladefranceschi">nicoladefranceschi</a> for this contribution. This change addresses <a href="https://github.com/exceljs/exceljs/issues/517">Ability to use external "moment" package #517</a>. </li> </ul> |
| 1.1.0   | <ul> <li> Addressed <a href="https://github.com/exceljs/exceljs/issues/514">Is there a way inserting values in columns. #514</a>. Added a new getter/setter property to Column to get and set column values (see <a href="#columns">Columns</a> for details). </li> </ul> |
| 1.1.1   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/532">Include index.d.ts in published packages #532</a>. To fix <a href="https://github.com/exceljs/exceljs/issues/525">TypeScript definitions missing from npm package #525</a>. Many thanks to <a href="https://github.com/saschanaz">Kagami Sascha Rosylight</a> for this contribution. </li> </ul> |
| 1.1.2   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/536">Don't break when docProps/core.xml contains <cp:contentType /> #536</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> (and reviewers) for this contribution. </li> </ul> |
| 1.1.3   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/537">Try to handle the case where a &lt;c&gt; element is missing an r attribute #537</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> </ul> |
| 1.2.0   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/544">Add dateUTC flag to CSV Writing #544</a>. Many thanks to <a href="https://github.com/zgriesinger">Zackery Griesinger</a> for this contribution. </li> </ul> |
| 1.2.1   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/547">worksheet name is writable #547</a>. Many thanks to <a href="https://github.com/f111fei">xzper</a> for this contribution. </li> </ul> |
| 1.3.0   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/549">Add CSV write buffer support #549</a>. Many thanks to <a href="https://github.com/jloveridge">Jarom Loveridge</a> for this contribution. </li> </ul> |
| 1.4.2   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/541">Discussion: Customizable row/cell limit #541</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> </ul> |
| 1.4.3   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/552">Get the right text out of hyperlinked formula cells #552</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> and <a href="https://github.com/holm">Christian Holm</a> for this contribution. </li> </ul> |
| 1.4.5   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/556">Add test case with a huge file #556</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> and <a href="https://github.com/holm">Christian Holm</a> for this contribution. </li> </ul> |
| 1.4.6   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/557">Update README.md to reflect correct functionality of row.addPageBreak() #557</a>. Many thanks to <a href="https://github.com/raj7desai">RajDesai</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/558">fix index.d.ts #558</a>. Many thanks to <a href="https://github.com/Diluka">Diluka</a> for this contribution. </li> </ul> |
| 1.4.7   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/562">List /xl/sharedStrings.xml in [Content_Types].xml only if one of the … #562</a>. Many thanks to <a href="https://github.com/priidikvaikla">Priidik Vaikla</a> for this contribution. </li> </ul> |
| 1.4.8   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/562">List /xl/sharedStrings.xml in [Content_Types].xml only if one of the … #562</a>. Many thanks to <a href="https://github.com/priidikvaikla">Priidik Vaikla</a> for this contribution. </li> <li> Fixed issue with above where shared strings were used but the content type was not added. </li> </ul> |
| 1.4.9   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/562">List /xl/sharedStrings.xml in [Content_Types].xml only if one of the … #562</a>. Many thanks to <a href="https://github.com/priidikvaikla">Priidik Vaikla</a> for this contribution. </li> <li> Fixed issue with above where shared strings were used but the content type was not added. </li> <li> Fixed issue <a href="https://github.com/exceljs/exceljs/issues/581">1.4.8 broke writing Excel files with useSharedStrings:true #581</a>. </li> </ul> |
| 1.4.10  | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/564">core-xform: Tolerate a missing cp: namespace for the coreProperties element #564</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> </ul> |
| 1.4.12  | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/567">Avoid error on malformed address #567</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/571">Added a missing Promise&lt;void&gt; in index.d.ts #571</a>. Many thanks to <a href="https://github.com/carboneater">Gabriel Fournier</a> for this contribution. This release should fix <a href="https://github.com/exceljs/exceljs/issues/548">Is workbook.commit() still a promise or not #548</a> </li> </ul> |
| 1.4.13  | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/574">Issue #488 #574</a>. Many thanks to <a href="https://github.com/dljenkins">dljenkins</a> for this contribution. This release should fix <a href="https://github.com/exceljs/exceljs/issues/488">Invalid time value Exception #488</a>. </li> </ul> |
| 1.5.0   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/577">Sheet add state for hidden or show #577</a>. Many thanks to <a href="https://github.com/Hsinfu">Freddie Hsinfu Huang</a> for this contribution. This release should fix <a href="https://github.com/exceljs/exceljs/issues/226">hide worksheet and reorder sheets #226</a>. </li> </ul> |
| 1.5.1   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/582">Update index.d.ts #582</a>. Many thanks to <a href="https://github.com/hankolsen">hankolsen</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/584">Decode the _x<4 hex chars>_ escape notation in shared strings #584</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> </ul> |
| 1.6.0   | <ul> <li> Added .html property to Cells to facilitate html-safe rendering. See <a href="#handling-individual-cells">Handling Individual Cells</a> for details. </li> </ul> |
| 1.6.1   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/587">Fix Issue #488 where dt is an invalid date format. #587</a> to fix  <a href="https://github.com/exceljs/exceljs/issues/488">Invalid time value Exception #488</a>. Many thanks to <a href="https://github.com/ilijaz">Iliya Zubakin</a> for this contribution. </li> </ul> |
| 1.6.2   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/587">Fix Issue #488 where dt is an invalid date format. #587</a> to fix  <a href="https://github.com/exceljs/exceljs/issues/488">Invalid time value Exception #488</a>. Many thanks to <a href="https://github.com/ilijaz">Iliya Zubakin</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/590">drawing element must be below rowBreaks according to spec or corrupt worksheet #590</a> Many thanks to <a href="https://github.com/nevace">Liam Neville</a> for this contribution. </li> </ul> |
| 1.6.3   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/595">set type optional #595</a> Many thanks to <a href="https://github.com/taoqf">taoqf</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/578">Fix some xlsx stream read xlsx not in guaranteed order problem #578</a> Many thanks to <a href="https://github.com/KMethod">KMethod</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/599">Fix formatting issue in README #599</a> Many thanks to <a href="https://github.com/getsomecoke">Vishnu Kyatannawar</a> for this contribution. </li> </ul> |
| 1.7.0   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/602">Ability to set tooltip for hyperlink #602</a> Many thanks to <a href="https://github.com/kalexey89">Kuznetsov Aleksey</a> for this contribution. </li> </ul> |
| 1.8.0   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/636">Fix misinterpreted ranges from &lt;definedName&gt; #636</a> Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/640">Add LGTM code quality badges #640</a> Many thanks to <a href="https://github.com/xcorail">Xavier RENE-CORAIL</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/646">Add type definition for Column.values #646</a> Many thanks to <a href="https://github.com/emlai">Emil Laine</a> for this contribution. This fixes <a href="https://github.com/exceljs/exceljs/issues/645">Column.values is missing TypeScript definitions #645</a>. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/663">Update README.md with load() option #663</a> Many thanks to <a href="https://github.com/thinksentient">Joanna Walker</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/677">fixed packages according to npm audit #677</a> Many thanks to <a href="https://github.com/misleadingTitle">Manuel Minuti</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/699">Update index.d.ts #699</a> Many thanks to <a href="https://github.com/rayyen">Ray Yen</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/708">Replaced node-unzip-2 to unzipper package which is more robust #708</a> Many thanks to <a href="https://github.com/johnmalkovich100">johnmalkovich100</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/728">Read worksheet hidden state #728</a> Many thanks to <a href="https://github.com/LesterLyu">Dishu(Lester) Lyu</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/736">add Worksheet.state typescript definition fix #714 #736</a> Many thanks to <a href="https://github.com/ilyes-kechidi">Ilyes Kechidi</a> for this contribution. This fixes <a href="https://github.com/exceljs/exceljs/issues/714">Worksheet State does not exist in index.d.ts #714</a>. </li> </ul> |
| 1.9.0   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/702">Improvements for images (correct reading/writing possitions) #702</a>. This fixes <a href="https://github.com/exceljs/exceljs/issues/650">Image location don't respect Column width #650</a> and <a href="https://github.com/exceljs/exceljs/issues/467">Image position - stretching image #467</a>. Many thanks to <a href="https://github.com/Siemienik">Siemienik Paweł</a> for this contribution. </li> </li> </ul> |
| 1.9.1   | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/619">Add Typescript support for formulas without results #619</a>. Many thanks to <a href="https://github.com/Wolfchin">Loursin</a> for this contribution. </li> </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/737">Fix existing row styles when using spliceRows #737</a>. Many thanks to <a href="https://github.com/cxam">cxam</a> for this contribution. </li> </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/774">Consistent code quality #774</a>. Many thanks to <a href="https://github.com/alubbe">Andreas Lubbe</a> for this contribution. </li> </li> </ul> |
| 1.10.0  | <ul> <li> Fixed effect of splicing rows and columns on defined names </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/746">Add support for adding images anchored to one cell #746</a>. Many thanks to <a href="https://github.com/karlvr">Karl von Randow</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/758">Add vertical align property #758</a>. Many thanks to <a href="https://github.com/MikeZyatkov">MikeZyatkov</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/775">Replace the temp lib to tmp #775</a>. Many thanks to <a href="https://github.com/coldhiber">Ivan Sotnikov</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/780">Replace the temp lib to tmp #775</a>. Many thanks to <a href="https://github.com/alubbe">Andreas Lubbe</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/793">Update Worksheet.dimensions return type #793</a>. Many thanks to <a href="https://github.com/Siemienik">Siemienik Paweł</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/795">One more types fix #795</a>. Many thanks to <a href="https://github.com/Siemienik">Siemienik Paweł</a> for this contribution. </li> </ul> |
| 1.11.0  | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/776">Add the ability to bail out of parsing if the number of columns exceeds a given limit #776</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/799">Add support for repeated columns on every page when printing. #799</a>. Many thanks to <a href="https://github.com/FreakenK">Jasmin Auger</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/815">Do not use a promise polyfill on modern setups #815</a>. Many thanks to <a href="https://github.com/alubbe">Andreas Lubbe</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/807">copy LICENSE to the dist folder #807</a>. Many thanks to <a href="https://github.com/zypA13510">Yuping Zuo</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/813">Avoid unhandled rejection on XML parse error #813</a>. Many thanks to <a href="https://github.com/papandreou">Andreas Lind</a> for this contribution. </li> </ul> |
| 1.12.0  | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/819">(chore) increment unzipper to 0.9.12 to address npm advisory 886 #819</a>. Many thanks to <a href="https://github.com/kreig303">Kreig Zimmerman</a> for this contribution. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/817">docs(README): improve docs #817</a>. Many thanks to <a href="https://github.com/zypA13510">Yuping Zuo</a> for this contribution. </li> <li> <p> Merged <a href="https://github.com/exceljs/exceljs/pull/823">add comment support #529 #823</a>. Many thanks to <a href="https://github.com/ilimei">ilimei</a> for this contribution. </p> <p>This fixes the following issues:</p> <ul> <li><a href="https://github.com/exceljs/exceljs/issues/202">Is it possible to add comment on a cell? #202</a></li> <li><a href="https://github.com/exceljs/exceljs/issues/451">Add comment to cell #451</a></li> <li><a href="https://github.com/exceljs/exceljs/issues/503">Excel add comment on cell #503</a></li> <li><a href="https://github.com/exceljs/exceljs/issues/529">How to add Cell comment #529</a></li> <li><a href="https://github.com/exceljs/exceljs/issues/707">Please add example to how I can insert comments for a cell #707</a></li> </ul> </li> </ul> |
| 1.12.1  | <ul> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/822">fix issue with print area defined name corrupting file #822</a>. Many thanks to <a href="https://github.com/donaldsonjulia">Julia Donaldson</a> for this contribution. This fixes issue <a href="https://github.com/exceljs/exceljs/issues/664">Defined Names Break/Corrupt Excel File into Repair Mode #664</a>. </li> <li> Merged <a href="https://github.com/exceljs/exceljs/pull/831">Only keep at most 31 characters for sheetname #831</a>. Many thanks to <a href="https://github.com/kaleo211">Xuebin He</a> for this contribution. This fixes issue <a href="https://github.com/exceljs/exceljs/issues/398">Limit worksheet name length to 31 characters #398</a>. </li> </ul> |

