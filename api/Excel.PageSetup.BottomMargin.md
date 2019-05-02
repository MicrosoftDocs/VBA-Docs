---
title: PageSetup.BottomMargin property (Excel)
keywords: vbaxl10.chm473074
f1_keywords:
- vbaxl10.chm473074
ms.prod: excel
api_name:
- Excel.PageSetup.BottomMargin
ms.assetid: 4c1cd3e0-0ba6-9d2d-4d5a-69d9ee811704
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.BottomMargin property (Excel)

Returns or sets the size of the bottom margin, in [points](../language/glossary/vbe-glossary.md#point). Read/write **Double**.


## Syntax

_expression_.**BottomMargin**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

Margins are set or returned in points. Use either the **[InchesToPoints](Excel.Application.InchesToPoints.md)** method or the **[CentimetersToPoints](Excel.Application.CentimetersToPoints.md)** method to do the conversion.


## Example

These two examples set the bottom margin of Sheet1 to 0.5 inch (36 points).

```vb
Worksheets("Sheet1").PageSetup.BottomMargin = _ 
 Application.InchesToPoints(0.5) 
 
Worksheets("Sheet1").PageSetup.BottomMargin = 36
```

<br/>

This example displays the current setting for the bottom margin on Sheet1.

```vb
marginInches = Worksheets("Sheet1").PageSetup.BottomMargin / _ 
 Application.InchesToPoints(1) 
MsgBox "The current bottom margin is " & marginInches & " inches"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]