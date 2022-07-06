# SpreadJS_FormulaFunction
公式函数

# SpreadJS_FormulaFunction
公式函数

### SpreadJS 示例，浮点数与公式
该示例包括使用 SpreadJS API 的演示脚本，可用于实现浮点数与公式
有关 SpreadJS API 的更多信息，请参阅[SpreadJS API指南]( https://demo.grapecity.com.cn/spreadjs/help/api/) 和[帮助手册]( https://help.grapecity.com.cn/pages/viewpage.action?pageId=5963808)。



### 运行步骤
1、在开始之前，请确保您已满足以下先决条件：
要运行 SpreadJS，浏览器必须支持 HTML5，客户端导入和导出 Excel 需要 IE10及以上。
请先了解 [SpreadJS 的产品使用环境]( https://www.grapecity.com.cn/developer/spreadjs/selection-guide/product-use-environment)，并申请临时部署授权激活
安装并更新NodeJS和NPM
2、克隆或下载此代码库
3、初始化控件，并运行示例脚本
#### 控件初始化
首先，创建一个新页面，并在页面上输入以下代码：
```
<!DOCTYPE html>
    <html>
    <head>
        <title>SpreadJS HTML Test Page</title>
```
2、在页面中添加对 SpreadJS 的引用。代码如下。需要注意的是，SpreadJS 提供压缩过
```
//（minified）的 JavaScript 文件和和用于调试的文件：
<script src="[Your_Scripts_Path]/gc.spread.sheets.all.xxxx.min.js" type="text/javascript"></script>
```
3、添加 CSS 文件以改变Spread.JS 的外观。默认的CSS文件名为： 
gc.spread.sheets.xxxx.css，里面包含了所有的默认样式。该 CSS 文件将会影响滚动条，筛选框及其子元素，单元格和下方标签栏的样式。引入 CSS 的代码如下：
```
//<link href="[Your_CSS_Path]/gc.spread.sheets.xxxx.css" rel="stylesheet" type="text/css"/>
```
4、添加产品授权，代码为（本地测试可以不添加）：
```
GC.Spread.Sheets.LicenseKey = "xxx";
```
5. 添加控件初始化代码。本例会在一个 id 为 “ss” 的 DOM 元素上初始化 SpreadJS：
```
<script type="text/javascript">
// Add your license
// If run this in local for testing, remove or comment below code
 GC.Spread.Sheets.LicenseKey = "xxx";

// Add your code
 window.onload = function(){
var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"),{sheetCount:3});
var sheet = spread.getActiveSheet();
 }
</script>
</head>
<body>
```
6、 创建一个 id 为 “ss” 的元素，SpreadJS 将在该 DOM 中初始化：
```
<div id="ss" style="height: 500px; width: 800px"></div>
</body>
</html>
```
#### 示例代码
```
HTML：
<div class="container">
     <div class="full-height clearfix mt2">
         <div class="col col-12 full-height ">
             <div id="ss" style="height:480px"></div>
         </div>
     </div>
 </div>

CSS：
  body {
      background: rgb(250, 250, 250);
      color: #333;
  }

  #ss {
      border: 1px #ccc solid;
  }

  .container {
      width: 80%;
      background: rgb(250, 250, 250);
      margin: 0 auto;
      height: 600px;
  }

  .full-height {
      height: 100%;
  }

  .left {
      height: 100%;
      overflow: auto;
  }

JavaScript：
// Title:浮点数与公式
// Description：浮点数与公式
// Tag:浮点数，公式
var spreadNS = GC.Spread.Sheets;


var toPrecisionFn = Number.prototype.toPrecision;
Number.prototype.toPrecision = function(precision) {
    precision = Math.min(11, precision);
    var number = toPrecisionFn.apply(this, arguments);
    return number;
}

var toFixFn = Number.prototype.toFixed;
Number.prototype.toFixed = function(precision) {
    precision = Math.min(11, precision);
    return toFixFn.apply(this, arguments);
}

$(document).ready(function() {
    var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));
    var sheet = spread.getActiveSheet();
    var menuData = spread.contextMenu.menuData;
    var newMenuData = [];
    menuData.forEach(function(item) {})
    spread.contextMenu.menuData = newMenuData;

    sheet.bind(GC.Spread.Sheets.Events.RangeChanged, function(sender, args) {
        console.log("RangeChanged");
        console.log(args);
    });

    /*----------------------------基本函数-----------------------*/
    sheet.suspendPaint();
    sheet.setValue(2, 1, '姓名');
    sheet.setValue(3, 1, '丁玉琴');
    sheet.setValue(4, 1, '杨国强');
    sheet.setValue(5, 1, '董超杨');
    sheet.setValue(6, 1, '杨猫猫');
    sheet.setValue(7, 1, '陈米灵');

    sheet.setValue(2, 2, '余额');
    sheet.setValue(3, 2, 342);
    sheet.setValue(4, 2, 3);
    sheet.setValue(5, 2, 5654);
    sheet.setValue(6, 2, 3455);
    sheet.setValue(7, 2, 2);

    sheet.setValue(9, 1, '平均：');
    sheet.setValue(10, 1, '合计：');
    sheet.setValue(11, 1, '最大值：');
    sheet.setValue(12, 1, '最小值：');
    sheet.setValue(13, 1, '名字包含杨的：');
    sheet.setValue(14, 1, '姓杨的：');

    sheet.setFormula(9, 2, '=AVERAGE(C4:C8)');
    sheet.setFormula(10, 2, '=SUM(C4:C8)');
    sheet.setFormula(11, 2, '=MAX(C4:C8)');
    sheet.setFormula(12, 2, '=MIN(C4:C8)');
    sheet.setFormula(13, 2, 'COUNTIF(B4:B8,"*杨*")');
    sheet.setFormula(14, 2, 'COUNTIF(B4:B8,"杨*")');
    sheet.setColumnWidth(1, 100)
    sheet.setValue(14, 5, '=MAX(C4:C7)');
    sheet.setFormula(15, 5, '=MAX(C4:C7)');

    sheet.addSpan(0, 0, 30, 1);
    sheet.setColumnCount(40);

    /*--------------------------INDIRECT函数----------------------*/

    sheet.setValue(1, 4, 234);
    sheet.setValue(2, 4, 'E2');
    sheet.setValue(3, 4, 'B4');
    sheet.setValue(4, 4, 23423);

    sheet.setValue(5, 5, 'INDIRECT("E1")=');
    sheet.setValue(6, 5, 'INDIRECT("B3")=');
    sheet.setValue(7, 5, 'INDIRECT("E"&(1+2))=');
    sheet.setValue(8, 5, 'INDIRECT(E4)=');
    sheet.setColumnWidth(5, 150);

    sheet.setFormula(5, 6, '=INDIRECT("E1")');
    sheet.setFormula(6, 6, '=INDIRECT("B3")');
    sheet.setFormula(7, 6, '=INDIRECT("E"&(1+2))');
    sheet.setFormula(8, 6, '=INDIRECT(E4)');

    sheet.resumePaint();
    /*-----------------------------自定义函数-------------------*/
    let sheet2 = new GC.Spread.Sheets.Worksheet();
    var spreadNS = GC.Spread.Sheets;
    spread.addSheet(1, sheet2)
    sheet2.setArray(1, 1, [
        ["序号", "底边长", "高", "面积"],
        [1, 4, 5],
        [2, 3, 4],
        [3, 1, 44],
        [4, 8, 3],
        [5, 4, 10],
        [6, 7, 10]
    ]);
    sheet2.addSpan(0, 1, 1, 4);
    sheet2.setValue(0, 1, "计算三角形面积");
    sheet2.getRange(0, 1, 1, 1).hAlign(spreadNS.HorizontalAlign.center);
    sheet2.setFormula(2, 4, '=(C3*D3)/2');
    sheet2.setValue(2, 0, '使用普通公式:');
    sheet2.setValue(3, 0, '使用自定义函数:');
    sheet2.setValue(7, 0, '异步函数:');
    sheet2.setValue(8, 0, '当前时间:');
    sheet2.setColumnWidth(0, 120);

    function calcuArea() {
        this.name = "area";
        this.maxArgs = 2;
        this.minArgs = 2;
    }
    calcuArea.prototype = new GC.Spread.CalcEngine.Functions.Function();
    calcuArea.prototype.evaluate = function(arg1, arg2) {
        if (arguments.length == 2 && !isNaN(parseInt(arg1)) && !isNaN(parseInt(arg2))) {
            return (arg1 * arg2) / 2;
        }
        return "#value"
    };
    var area = new calcuArea();
    sheet2.addCustomFunction(area);
    sheet2.setFormula(3, 4, "=area(C4,D4)");

    /*-----------------------数组公式----------------------*/
    sheet2.setValue(4, 0, '使用数组公式:');
    sheet2.addSpan(4, 0, 3, 1);
    sheet2.setArrayFormula(4, 4, 3, 1, "=(C5:C7*D5:D7)/2");

    /*----------------------异步函数---------------------*/
    var asyncSum = function() {
        this.name = "asyncArea";
        this.maxArgs = 2;
        this.minArgs = 2;
    };
    asyncSum.prototype = new GC.Spread.CalcEngine.Functions.AsyncFunction("ASUM", 1, 10);
    asyncSum.prototype.defaultValue = function() {
        return "计算中...";
    };
    asyncSum.prototype.evaluateAsync = function(context) {
        var args = arguments;

        var result = 0;
        setTimeout(function() {
            result = (args[1] * args[2]) / 2;
            console.log(args[1]);
            console.log(args[2]);
            context.setAsyncResult(result);
        }, 3000);
    };
    var asyncTime = function() {
        this.name = "asyncTime";
        this.maxArgs = 2;
        this.minArgs = 0;
    };
    asyncTime.prototype = new GC.Spread.CalcEngine.Functions.AsyncFunction("ASUM", 1, 10);
    asyncTime.prototype.evaluateAsync = function(context) {
        var args = arguments;
        var time = new Date().toString();
        context.setAsyncResult(time);
    };
    var asyncArea = new asyncSum();
    sheet2.addCustomFunction(asyncArea);
    sheet2.setFormula(7, 4, "=asyncArea(C8,D8)");
    var asyncTime = new asyncTime();
    sheet2.addCustomFunction(asyncTime);
    setInterval(function() {
        sheet2.setFormula(8, 1, "=asyncTime()");
    }, 1000);
});
```

#### 关于 SpreadJS
[SpreadJS]( https://www.grapecity.com.cn/developer/spreadjs) 是一款基于 HTML5 的纯前端表格控件，兼容 450 多种 Excel 公式，具备“高性能、跨平台、与 Excel 高度兼容”的产品特性。使用 SpreadJS，可直接在 Angular、 React、 Vue 等前端框架中实现高效的模板设计、在线编辑和数据绑定等功能，为最终用户提供高度类似 Excel 的使用体验。

