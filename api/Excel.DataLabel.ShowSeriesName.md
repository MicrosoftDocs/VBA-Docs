---
title: DataLabel.ShowSeriesName property (Excel)
keywords: vbaxl10.chm582099
f1_keywords:
- vbaxl10.chm582099
api_name:
- Excel.DataLabel.ShowSeriesName
ms.assetid: 95fd3b99-1ea5-5b51-7048-1dfba228aaa6
ms.date: 04/23/2019
ms.localizationpriority: medium
---


# DataLabel.ShowSeriesName property (Excel)

Returns or sets a **Boolean** to indicate the series name display behavior for the data labels on a chart. **True** to show the series name. **False** to hide. Read/write.


## Syntax

_expression_.**ShowSeriesName**

_expression_ A variable that represents a **[DataLabel](excel.datalabel(object).md)** object.


## Remarks

The chart must first be active before you can access the data labels programmatically, or a run-time error will occur.


## Example

This example enables the series name to be shown for the data labels of the first series on the first chart. This example assumes that a chart exists on the active worksheet.

```vb
Sub UseSeriesName() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowSeriesName = True 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]