---
title: ShowPercentage property (Excel Graph)
keywords: vbagr10.chm3077090
f1_keywords:
- vbagr10.chm3077090
api_name:
- Excel.ShowPercentage
ms.assetid: 32e2e547-8fb6-f3c7-3f61-a32a5d77d98d
ms.date: 04/12/2019
ms.localizationpriority: medium
---


# ShowPercentage property (Excel Graph)

Allows the user to show the percentage value for the data labels on a chart. Read/write **Boolean**.

## Syntax

_expression_.**ShowPercentage**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

The chart must first be active before you can access the data labels programmatically.


## Example

This example enables the percentage value to be shown for the data labels of the first series on the first chart.

```vb
Sub UsePercentage() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowPercentage = True 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]