---
title: ShowCategoryName property (Excel Graph)
keywords: vbagr10.chm67467
f1_keywords:
- vbagr10.chm67467
ms.prod: excel
api_name:
- Excel.ShowCategoryName
ms.assetid: f66a0162-f1b7-5b8d-ae09-bb928751cde3
ms.date: 04/12/2019
localization_priority: Normal
---


# ShowCategoryName property (Excel Graph)

Allows the user to show the category name for the data labels on a chart. Read/write **Boolean**.

## Syntax

_expression_.**ShowCategoryName**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

The chart must first be active before you can access the data labels programmatically.


## Example

This example enables the category name to be shown for the data labels of the first series on the first chart.

```vb
Sub UseCategoryName() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowCategoryName = True 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]