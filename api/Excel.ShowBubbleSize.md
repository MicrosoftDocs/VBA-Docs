---
title: ShowBubbleSize property (Excel Graph)
keywords: vbagr10.chm3077088
f1_keywords:
- vbagr10.chm3077088
ms.prod: excel
api_name:
- Excel.ShowBubbleSize
ms.assetid: abeee041-0fa0-537e-6786-37213a6004c1
ms.date: 04/12/2019
localization_priority: Normal
---


# ShowBubbleSize property (Excel Graph)

Allows the user to show the bubble size for the data labels on a chart. Read/write **Boolean**.

## Syntax

_expression_.**ShowBubbleSize**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

The chart must first be active before you can access the data labels programmatically.


## Example

This example enables the bubble size to be shown for the data labels of the first series on the first chart.

```vb
Sub UseBubbleSize() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowBubbleSize = True 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]