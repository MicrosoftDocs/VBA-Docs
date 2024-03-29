---
title: ShowValue property (Excel Graph)
keywords: vbagr10.chm67468
f1_keywords:
- vbagr10.chm67468
api_name:
- Excel.ShowValue
ms.assetid: 43e4380c-8e28-627e-6211-f1bd96d9d47f
ms.date: 04/12/2019
ms.localizationpriority: medium
---


# ShowValue property (Excel Graph)

Allows the user to show the value for the data labels on a chart. Read/write **Boolean**.

## Syntax

_expression_.**ShowValue**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

The chart must first be active before you can access the data labels programmatically.


## Example

This example enables the value to be shown for the data labels of the first series on the first chart.

```vb
Sub UseValue() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowValue = True 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]