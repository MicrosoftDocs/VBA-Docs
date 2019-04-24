---
title: DataLabel.ShowValue property (Excel)
keywords: vbaxl10.chm582101
f1_keywords:
- vbaxl10.chm582101
ms.prod: excel
api_name:
- Excel.DataLabel.ShowValue
ms.assetid: 83d4ead9-3539-d420-d4bd-2b474e174e10
ms.date: 04/23/2019
localization_priority: Normal
---


# DataLabel.ShowValue property (Excel)

Returns or sets a **Boolean** corresponding to a specified chart's data label values display behavior. **True** displays the values. **False** to hide. Read/write.


## Syntax

_expression_.**ShowValue**

_expression_ A variable that represents a **[DataLabel](excel.datalabel(object).md)** object.


## Remarks

The specified chart must first be active before you can access the data labels programmatically, or a run-time error will occur.


## Example

This example enables the value to be shown for the data labels of the first series, on the first chart. This example assumes that a chart exists on the active worksheet.

```vb
Sub UseValue() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowValue = True 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]