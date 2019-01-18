---
title: Worksheet.VPageBreaks property (Excel)
keywords: vbaxl10.chm175136
f1_keywords:
- vbaxl10.chm175136
ms.prod: excel
api_name:
- Excel.Worksheet.VPageBreaks
ms.assetid: 2a8d5c77-a609-4995-7216-de71295eda9a
ms.date: 06/08/2017
localization_priority: Normal
---


# Worksheet.VPageBreaks property (Excel)

Returns a  **[VPageBreaks](Excel.Worksheet.VPageBreaks.md)** collection that represents the vertical page breaks on the sheet. Read-only.


## Syntax

_expression_. `VPageBreaks`

_expression_ A variable that represents a [Worksheet](./Excel.Worksheet.md) object.


## Example

This example displays the total number of full-screen and print-area vertical page breaks.


```vb
For Each pb in Worksheets(1).VPageBreaks 
 If pb.Extent = xlPageBreakFull Then 
 cFull = cFull + 1 
 Else 
 cPartial = cPartial + 1 
 End If 
Next 
MsgBox cFull & " full-screen page breaks, " & cPartial & _ 
 " print-area page breaks"
```


## See also


[Worksheet Object](Excel.Worksheet.md)

