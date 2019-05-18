---
title: Worksheets.HPageBreaks property (Excel)
keywords: vbaxl10.chm470084
f1_keywords:
- vbaxl10.chm470084
ms.prod: excel
api_name:
- Excel.Worksheets.HPageBreaks
ms.assetid: d5541a3f-df09-a8cf-8a40-90a014b0c464
ms.date: 05/18/2019
localization_priority: Normal
---


# Worksheets.HPageBreaks property (Excel)

Returns an **[HPageBreaks](Excel.HPageBreaks.md)** collection that represents the horizontal page breaks on the sheet. Read-only.


## Syntax

_expression_.**HPageBreaks**

_expression_ A variable that represents a **[Worksheets](Excel.Worksheets.md)** object.


## Remarks

There is a limit of 1026 horizontal page breaks per sheet.


## Example

This example displays the number of full-screen and print-area horizontal page breaks.

```vb
For Each pb in Worksheets(1).HPageBreaks 
 If pb.Extent = xlPageBreakFull Then 
 cFull = cFull + 1 
 Else 
 cPartial = cPartial + 1 
 End If 
Next 
MsgBox cFull & " full-screen page breaks, " & cPartial & _ 
 " print-area page breaks"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]