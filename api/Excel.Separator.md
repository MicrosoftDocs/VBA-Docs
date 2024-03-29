---
title: Separator property (Excel Graph)
keywords: vbagr10.chm3077086
f1_keywords:
- vbagr10.chm3077086
api_name:
- Excel.Separator
ms.assetid: d83c68fc-5948-a65f-b3bb-09e3a3884163
ms.date: 04/12/2019
ms.localizationpriority: medium
---


# Separator property (Excel Graph)

Allows the user to set or return the separator used for the data labels on a chart. Read/write **Variant**.

## Syntax

_expression_.**Separator**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

The chart must first be active before you can access the data labels programmatically.


## Example

This example sets the data label separator for the first series on the first chart to a semi-colon.

```vb
Sub ChangeSeparator() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.Separator = ";" 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]