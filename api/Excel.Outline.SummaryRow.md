---
title: Outline.SummaryRow property (Excel)
keywords: vbaxl10.chm455076
f1_keywords:
- vbaxl10.chm455076
ms.prod: excel
api_name:
- Excel.Outline.SummaryRow
ms.assetid: f36fac55-cafd-1ec6-4e85-a7f4fc665c04
ms.date: 05/02/2019
localization_priority: Normal
---


# Outline.SummaryRow property (Excel)

Returns or sets the location of the summary rows in the outline. Read/write **[XlSummaryRow](Excel.XlSummaryRow.md)**.


## Syntax

_expression_.**SummaryRow**

_expression_ A variable that represents an **[Outline](Excel.Outline.md)** object.


## Remarks

Set **XlSummaryRow** to **xlAbove** for Microsoft Word-style outlines, where category headers are above the detailed information. 

Set **XlSummaryRow** to **xlBelow** for accounting-style outlines, where summations are below the detailed information.


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