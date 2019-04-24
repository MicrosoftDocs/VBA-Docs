---
title: DataLabels.ShowBubbleSize property (Excel)
keywords: vbaxl10.chm584103
f1_keywords:
- vbaxl10.chm584103
ms.prod: excel
api_name:
- Excel.DataLabels.ShowBubbleSize
ms.assetid: b7fe576f-c736-4e64-1c24-ec21273e237f
ms.date: 04/23/2019
localization_priority: Normal
---


# DataLabels.ShowBubbleSize property (Excel)

**True** to show the bubble size for the data labels on a chart. **False** to hide. Read/write **Boolean**.


## Syntax

_expression_.**ShowBubbleSize**

_expression_ An expression that returns a **[DataLabels](Excel.DataLabels(object).md)** object.


## Remarks

The chart must first be active before you can access the data labels programmatically, or a run-time error will occur.


## Example

This example shows the bubble size for the data labels of the first series on the first chart. This example assumes that a chart exists on the active worksheet.

```vb
Sub UseBubbleSize() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowBubbleSize = True 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]