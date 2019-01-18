---
title: HPageBreak.Extent property (Excel)
keywords: vbaxl10.chm159077
f1_keywords:
- vbaxl10.chm159077
ms.prod: excel
api_name:
- Excel.HPageBreak.Extent
ms.assetid: 07dc69ce-f46e-0b0d-412c-d22a9dbf5050
ms.date: 06/08/2017
localization_priority: Normal
---


# HPageBreak.Extent property (Excel)

Returns the type of the specified page break: full-screen or only within a print area. Can be either of the following  **[xlPageBreakExtent](Excel.XlPageBreakExtent.md)** constants: **xlPageBreakFull** or **xlPageBreakPartial**. Read-only **Long**.


## Syntax

_expression_. `Extent`

_expression_ A variable that represents a [HPageBreak](Excel.HPageBreak.md) object.


## Example

This example displays the total number of full-screen and print-area horizontal page breaks.


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


## See also


[HPageBreak Object](Excel.HPageBreak.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]