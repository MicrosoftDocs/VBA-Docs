---
title: Position property (Excel Graph)
keywords: vbagr10.chm65669
f1_keywords:
- vbagr10.chm65669
ms.prod: excel
api_name:
- Excel.Position
ms.assetid: 0e9e41e2-30a8-c744-72d1-3820cc4975f2
ms.date: 04/11/2019
localization_priority: Normal
---


# Position property (Excel Graph)

The **Position** property as it applies to the **DataLabel**, **DataLabels**, and **Legend** objects.

## DataLabel and DataLabels objects

Returns or sets the position of the data label. Read/write **[XlDataLabelPosition](excel.xldatalabelposition.md)**.

### Syntax

_expression_.**Position**

_expression_ Required. An expression that returns a **[DataLabel](excel.datalabel-graph-object.md)** object or **[DataLabels](excel.datalabels(collection).md)** collection.




## Legend object

Returns or sets the position of the legend on the chart. Read/write **[XlLegendPosition](excel.xllegendposition.md)**.

### Syntax

_expression_.**Position**

_expression_ Required. An expression that returns a **[Legend](excel.legend-graph-object.md)** object.

### Example

This example sets the position of the legend to the top of the chart.

```vb
myChart.Legend.Position = xlLegendPositionTop
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]