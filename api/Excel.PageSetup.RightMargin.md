---
title: PageSetup.RightMargin property (Excel)
keywords: vbaxl10.chm473101
f1_keywords:
- vbaxl10.chm473101
ms.prod: excel
api_name:
- Excel.PageSetup.RightMargin
ms.assetid: 9c392522-2a06-c76f-2f7a-0fb93c947d39
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.RightMargin property (Excel)

Returns or sets the size of the right margin, in points. Read/write  **Double**.


## Syntax

_expression_.**RightMargin**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

Margins are set or returned in points. Use the  **InchesToPoints** method or the **CentimetersToPoints** method to convert measurements from inches or centimeters.


## Example

This example sets the right margin of Sheet1 to 1.5 inches.


```vb
Worksheets("Sheet1").PageSetup.RightMargin = _ 
 Application.InchesToPoints(1.5)
```

This example sets the right margin of Sheet1 to 2 centimeters.




```vb
Worksheets("Sheet1").PageSetup.RightMargin = _ 
 Application.CentimetersToPoints(2)
```

This example displays the current right-margin setting for Sheet1.




```vb
marginInches = Worksheets("Sheet1").PageSetup.RightMargin / _ 
 Application.InchesToPoints(1) 
MsgBox "The current right margin is " & marginInches & " inches"
```


## See also


[PageSetup Object](Excel.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]