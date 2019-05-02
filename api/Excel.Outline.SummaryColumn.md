---
title: Outline.SummaryColumn property (Excel)
keywords: vbaxl10.chm455075
f1_keywords:
- vbaxl10.chm455075
ms.prod: excel
api_name:
- Excel.Outline.SummaryColumn
ms.assetid: b134c991-7875-445a-ca73-d48bf23f3eea
ms.date: 05/02/2019
localization_priority: Normal
---


# Outline.SummaryColumn property (Excel)

Returns or sets the location of the summary columns in the outline. Read/write **[XlSummaryColumn](Excel.XlSummaryColumn.md)**.


## Syntax

_expression_.**SummaryColumn**

_expression_ A variable that represents an **[Outline](Excel.Outline.md)** object.


## Example

This example creates an outline with automatic styles, with the summary row above the detail rows, and with the summary column to the right of the detail columns.

```vb
Worksheets("Sheet1").Activate 
Selection.AutoOutline 
With ActiveSheet.Outline 
 .SummaryRow = xlAbove 
 .SummaryColumn = xlRight 
 .AutomaticStyles = True 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]