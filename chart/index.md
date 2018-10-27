# Office Add-in数据可视化的最佳实践

## 简介
Office Add-in平台允许您构建扩展 Office 应用程序并与Office文档中的内容进行交互的解决方案。通过Office Add-in可以使用熟悉的 Web 技术，例如 HTML、CSS 和 JavaScript 来扩展 Word、Excel、PowerPoint、OneNote、Project 和 Outlook，并与其进行交互。在本课程中，您将使用我们最新的Excel API对图表进行深度定制，帮助您的应用更好的向用户展示数据。

### 动手实验的目标
这个实验将会向您展示
- 如何使用Script Lab
- 如何使用Excel Chart API

### 系统需求
您需要安装
- Windows 10
- Office 365或Office 2016/2019

### 配置
您需要执行以下步骤来为这次实验准备环境
- 安装Microsoft Windows 10.
- 安装Microsoft Office 365或者Microsoft Office 2016/2019

### 练习
本次动手实验包含以下几个部分:

1\. 使用JS API在Excel里创建一个Chart

2\. 对Chart做深度定制化

3\. 具体的一个数据可视化实践

## Exercise 1: 使用JS API在Excel里创建一个Chart


### Task 1 – 启动Script Lab
在我们开始前，我们需要启动Excel并安装Script Lab

1\. 在Insert tab里点击Get Add-ins . 
![动手实验1](images/01.png?raw=true)

2\. 在Office Add-in Store里搜索Script Lab并点击Add.
![动手实验1](images/02.png?raw=true)

3\. 等待片刻，Script Lab就会出现在Ribbon里，点击Code以启动代码编辑器，点击Run以启动代码运行界面
![动手实验1](images/03.png?raw=true)

4\. 代码编辑器和运行界面会以task pane的方式在侧边栏出现
![动手实验1](images/04.png?raw=true)
 
### Task 2 – 加入样本数据
我们先下载用于练习的Excel文件([下载链接](samples/exercise.xlsx))，并打开“Exercise 1&2”工作表

![动手实验1](images/05.png?raw=true)

### Task 3 – 通过JavaScript API 创建一个Chart
在代码编辑器中键入以下代码：

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("A1:B11");
    let chart = sheet.charts.add(Excel.ChartType.lineMarkers, range);
    chart.title.text = "历史温度";

    await context.sync();
})
```

在运行界面点击Run button来运行刚刚所写代码,得到结果如下：
![动手实验1](images/06.png?raw=true)
至此，您就成功的通过JS API在Excel里添加了一个Chart。

## Exercise 2: 对Chart做深度定制化
这个练习主要是通过一些简单的场景来展示如何通过API在一个Chart里添加、删除、定制series、trendline、title以及datalabel。您可以使用上个练习得到的Chart继续下面的环节。

### Task 1 – 添加、删除及定制Series 
在最新发布的API中，我们对Chart的Series系列API做了进一步的扩展，使得开发者可以对其做更多的定制。其中值得说明的是我们允许您对已有的Chart添加、删除series，这样您在做数据可视化时可以更灵活的引用Excel里不同的区域的数据。下面我们就以这个特性简单展示一下Series相关的API。

1\. 在Script Lab里键入如下代码：

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.getItemAt(0);
    let valueRange = sheet.getRange("C2:C11");
    let categoryRange = sheet.getRange("A2:A11");
    let series = chart.series.add();
    series.setXAxisValues(categoryRange);
    series.setValues(valueRange);

    await context.sync();
})
```

2\. 在Script Lab里点击”Run”,启动运行，我们会在chart里添加一条新的series来展示10月1日至10日的最低温度，如下图：
 ![动手实验2](images/07.png?raw=true)

3\. 上面的步骤是在一个chart里添加了一个series，我们通过下面的代码可以改变series的样式，从而更清楚地展示每日最高温度。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    series.markerStyle = Excel.ChartMarkerStyle.diamond;
    series.markerBackgroundColor = 'red';
    series.hasDataLabels = true;

    await context.sync();
})
```

4\. 运行代码得到如下结果：
 ![动手实验2](images/08.png?raw=true)

5\. 如果我们这时不关心最高温度了，那么我们可以通过如下代码将其从Chart里删除

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    series.delete();

    await context.sync();
})
```

6\. 点击运行按键则可得到如您所需的Chart
 ![动手实验2](images/09.png?raw=true)
 
### Task 2 – 添加/删除Trendline
同时我们在新的API里加入了对Trendline的支持。现在您可以在Chart里对任意一个series添加或删除一个或多个Trendline。

1\. 在代码编辑器键入以下代码
```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let trendline = series.trendlines.add(Excel.ChartTrendlineType.polynomial);

    await context.sync();
})
```

2\. 这样我们就很轻松的对最低气温添加了一个Trendline，可以用来展示数据的运行趋势，如下图：
  ![动手实验2](images/10.png?raw=true)

3\. 我们还可以对这个Trendline做一些样式上的定制

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let trendline = series.trendlines.getItem(0);
    trendline.showEquation = true;
    trendline.format.line.color = 'green';
    trendline.format.line.lineStyle = Excel.ChartLineStyle.dashDotDot;

    await context.sync();
})
```

4\. 在上面的代码里，我们将trendline的颜色设置为绿色，把trendline的计算公式显示出来，并将趋势线的样式稍微改变了一下，结果如下图
  ![动手实验2](images/11.png?raw=true)

5\. 同样的，如果不再需要这个trendline，我们也可以通过API将其删除

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let trendline = series.trendlines.getItem(0);
    trendline.delete();

    await context.sync();
})
```

### Task 3 – 定制Chart Title和DataLabel
在新发布的Chart API里，我们还对Chart的Title和DataLabel提供了更多的定制能力。比如原来我们只可以对整个Title设置字体属性，但是现在我们可以对其部分文字进行设置。这样开发者在给用户提供一个chart的时候，可以高亮关键词。

1\.	我们先来对Chart的title做一些改变。在代码编辑器键入以下代码

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.getItemAt(0);
    chart.title.text = "上海市十月上旬最低温度";
    let textrange = chart.title.getSubstring(0,3);
    let font = textrange.font;
    font.size = 18;
    font.color = 'red';
    font.bold = true;
    font.italic = true;

    await context.sync();
})
```
在代码中，我们将chart的title设置为“上海市十月上旬最低温度“，对“上海市”三个字做了各种字体属性的设置以高亮，结果如下图：
 ![动手实验2](images/12.png?raw=true)

2\.	我们还可以对数据中的每一个单独的datalabel做定制，键入如下代码：

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let point = series.points.getItemAt(3); // 第四个数据点
    point.hasDataLabel = true;
    point.markerSize = 8;
    point.markerBackgroundColor = 'red';
    point.markerStyle = Excel.ChartMarkerStyle.diamond;
    let datalabel = point.dataLabel;
    datalabel.showCategoryName = true;
    datalabel.showValue = true;
    datalabel.showLegendKey = true;
    datalabel.position = Excel.ChartDataLabelPosition.left;

    await context.sync();
})
```
在上面的代码中，我们将series里温度最高的第四天，即十月四日，所对应的数据点和datalabel高亮
![动手实验2](images/13.png?raw=true)
 
## Exercise 3: 具体的一个数据可视化实践
在这个练习里，我们将为大家提供一组示例数据。大家可以使用本课程中学到的API构建一个Chart。

下面为样本数据([下载链接](samples/exercise.xlsx))：

![动手实验3](images/14.png?raw=true)

最终的Chart如下

![动手实验3](images/15.png?raw=true)

其中我们需要：

1\. 添加一个空白的line chart

2\. 添加一个series，其数值为收入

3\. 将chart的标题设置为“七月份销售趋势图”，并高亮七月份

4\. 为series添加一个polynomial趋势线

5\. 设置category axis和value axis的区间

6\. 高亮最高销售和最低销售数据点并添加datalabel


大家可以尝试自己写一下这个代码。下面提供一些有用的代码段：
- 得到练习3的工作表
```js
let sheet = context.workbook.worksheets.getItem("Exercise 3");
```
- 得到日期列和收入列的Range
```js
let table = context.workbook.tables.getItem("SalesTable");
let salesRange = table.columns.getItem("收入").getDataBodyRange();
let dateRange = table.columns.getItem("日期").getDataBodyRange();
```
- 设置value axis的displayunit
```js
let valueAxis = chart.axes.valueAxis;
valueAxis.displayUnit = Excel.ChartAxisDisplayUnit.thousands;
```
大家练习完可以对照我给出的答案作为参考([链接](https://gist.github.com/techsummit2018/a5234d954d177ca8fb521ec6fea44c94))

经过上面一系列的动手练习，相信大家已经基本掌握如何使用Script Lab来使用Chart API。在Office JS API 1.7和1.8版本，我们加入了数百个新的Chart API供大家使用，具体的列表请访问下面的网页做进一步的研究。

https://docs.microsoft.com/en-us/javascript/api/excel?view=office-js 

希望我们新的API可以帮助诸位为客户创造更多的解决方案，谢谢！
