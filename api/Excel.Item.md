---
title: Item method (Excel Graph)
keywords: vbagr10.chm3077621
f1_keywords:
- vbagr10.chm3077621
ms.prod: excel
api_name:
- Excel.Item
ms.assetid: 9e92de7f-b231-c7c5-fcea-50c1051d1add
ms.date: 04/09/2019
localization_priority: Normal
---


# Item method (Excel Graph)

The **Item** method as it applies to the **Axes**, **ChartGroups**, **DataLabels**, **LegendEntries**, **Points**, **SeriesCollection**, and **Trendlines** collections.

## Axes collection

Returns a single **Axis** object from an **Axes** collection.

### Syntax

_expression_.**Item** (_Type_, _AxisGroup_)

_expression_ Required. An expression that returns an **[Axes](excel.axes-graph-collection.md)** collection.

### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Type_ |Required |**[XlAxisType](excel.xlaxistype.md)**|The axis type. Can be one of these **XlAxisType** constants: **xlCategory**, **xlSeriesAxis** (valid only for 3D charts), or **xlValue**.|
|_AxisGroup_ |Optional |**[XlAxisGroup](excel.xlaxisgroup.md)**|The axis group. Can be one of these **XlAxisGroup** constants: **xlSecondary** or **xlPrimary** (default). |

### Example

This example sets the title text for the category axis on Chart1.

```vb
With Charts("chart1").Axes.Item(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```


## ChartGroups collection

Returns a single **ChartGroup** object from a **ChartGroups** collection.

### Syntax

_expression_.**Item** (_Index_)

_expression_ Required. An expression that returns a **[ChartGroups](excel.chartgroups(collection).md)** collection.

### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Required |**Variant** |The index number of the chart group.|

### Example

This example adds drop lines to chart group one on chart sheet one.

```vb
Charts(1).ChartGroups.Item(1).HasDropLines = True
```


## DataLabels collection

Returns a single **DataLabel** object from a **DataLabels** collection.

### Syntax

_expression_.**Item** (_Index_)

_expression_ Required. An expression that returns a **[DataLabels](excel.datalabels(collection).md)** collection.

### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Required |**Variant**|The name or index number of the data label.|

### Example

This example sets the number format for the fifth data label in series one in embedded chart one on worksheet one.

```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).DataLabels.Item(5).NumberFormat = "0.000"
```



## LegendEntries collection

Returns a single **LegendEntry** object from a **LegendEntries** collection.

### Syntax

_expression_.**Item** (_Index_)

_expression_ Required. An expression that returns a **[LegendEntries](excel.legendentries(collection).md)** collection.

### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Required |**Variant**|The index number of the legend entry.|

### Example

This example changes the font for the text of the legend entry at the top of the legend (this is usually the legend for series one) in embedded chart one on Sheet1.

```vb
Worksheets("sheet1").ChartObjects(1).Chart _ 
 .Legend.LegendEntries.Item(1).Font.Italic = True
```




## Points collection

Returns a single **Point** object from a **Points** collection.

### Syntax

_expression_.**Item** (_Index_)

_expression_ Required. An expression that returns a **[Points](excel.points(collection).md)** collection.

### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Required |**Long**|The index number of the point.|

### Example

This example sets the marker style for the third point in series one in embedded chart one on worksheet one. The specified series must be a 2D line, scatter, or radar series.

```vb
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection(1).Points.Item(3).MarkerStyle = xlDiamond
```




## SeriesCollection collection

Returns a single **Series** object from a **SeriesCollection** collection.

### Syntax

_expression_.**Item** (_Index_)

_expression_ Required. An expression that returns a **[SeriesCollection](excel.seriescollection(collection).md)** collection.

### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Required |**Variant**|The name or index number of the series.|

### Example

This example provides two lines of code that are equivalent.

```vb
myChart.SeriesCollection.Item(1) 
myChart.SeriesCollection(1)
```


## Trendlines collection

Returns a single **Trendline** object from a **Trendlines** collection.

### Syntax

_expression_.**Item** (_Index_)

_expression_ Required. An expression that returns a **[Trendlines](excel.trendlines(collection).md)** collection.

### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Optional |**Variant**|The name or index number of the trendline.|

### Example

This example sets the number of units that the trendline on Chart1 extends forward and backward. The example should be run on a 2D column chart that contains a single series with a trendline.

```vb
With Charts("Chart1").SeriesCollection(1).Trendlines.Item(1) 
 .Forward = 5 
 .Backward = .5 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]