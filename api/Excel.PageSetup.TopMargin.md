---
title: PageSetup.TopMargin property (Excel)
keywords: vbaxl10.chm473102
f1_keywords:
- vbaxl10.chm473102
ms.prod: excel
api_name:
- Excel.PageSetup.TopMargin
ms.assetid: 1c4efb20-844c-b602-48b4-ef60b8e5dda5
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.TopMargin property (Excel)

Returns or sets the size of the top margin, in points. Read/write  **Double**.


## Syntax

_expression_.**TopMargin**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

Margins are set or returned in points. Use the  **InchesToPoints** method or the **CentimetersToPoints** method to convert measurements from inches or centimeters.


## Example

These two examples set the top margin of Sheet1 to 0.5 inch (36 points).


```vb
Worksheets("Sheet1").PageSetup.TopMargin = _ 
 Application.InchesToPoints(0.5) 
 
Worksheets("Sheet1").PageSetup.TopMargin = 36
```

This example displays the current top-margin setting.




```vb
marginInches = ActiveSheet.PageSetup.TopMargin / _ 
 Application.InchesToPoints(1) 
MsgBox "The current top margin is " & marginInches & " inches"
```


## See also


[PageSetup Object](Excel.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]