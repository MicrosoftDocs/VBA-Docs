---
title: PageSetup.LeftMargin property (Excel)
keywords: vbaxl10.chm473088
f1_keywords:
- vbaxl10.chm473088
ms.prod: excel
api_name:
- Excel.PageSetup.LeftMargin
ms.assetid: 5d52ca64-6fe7-5c0e-63ab-036aa5119bb2
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.LeftMargin property (Excel)

Returns or sets the size of the left margin, in [points](../language/glossary/vbe-glossary.md#point). Read/write **Double**.


## Syntax

_expression_.**LeftMargin**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

Margins are set or returned in points. Use either the **[InchesToPoints](Excel.Application.InchesToPoints.md)** method or the **[CentimetersToPoints](Excel.Application.CentimetersToPoints.md)** method to do the conversion.


## Example

This example sets the left margin of Sheet1 to 1.5 inches.

```vb
Worksheets("Sheet1").PageSetup.LeftMargin = _ 
 Application.InchesToPoints(1.5)
```

<br/>

This example sets the left margin of Sheet1 to 2 centimeters.

```vb
Worksheets("Sheet1").PageSetup.LeftMargin = _ 
 Application.CentimetersToPoints(2)
```

<br/>

This example displays the current left-margin setting for Sheet1.

```vb
marginInches = Worksheets("Sheet1").PageSetup.LeftMargin / _ 
 Application.InchesToPoints(1) 
MsgBox "The current left margin is " & marginInches & " inches"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
