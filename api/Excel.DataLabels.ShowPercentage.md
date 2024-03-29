---
title: DataLabels.ShowPercentage property (Excel)
keywords: vbaxl10.chm584102
f1_keywords:
- vbaxl10.chm584102
api_name:
- Excel.DataLabels.ShowPercentage
ms.assetid: c8afd00d-3443-8366-6c74-d426237c6fd7
ms.date: 04/23/2019
ms.localizationpriority: medium
---


# DataLabels.ShowPercentage property (Excel)

**True** to display the percentage value for the data labels on a chart. **False** to hide. Read/write **Boolean**.


## Syntax

_expression_.**ShowPercentage**

_expression_ A variable that represents a **[DataLabels](Excel.DataLabels(object).md)** object.


## Remarks

The chart must first be active before you can access the data labels programmatically, or a run-time error will occur.


## Example

This example enables the percentage value to be shown for the data labels of the first series on the first chart. This example assumes that a chart exists on the active worksheet.

```vb
Sub UsePercentage() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowPercentage = True 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]