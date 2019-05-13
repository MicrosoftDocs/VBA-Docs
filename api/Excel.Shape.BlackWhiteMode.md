---
title: Shape.BlackWhiteMode property (Excel)
keywords: vbaxl10.chm636118
f1_keywords:
- vbaxl10.chm636118
ms.prod: excel
api_name:
- Excel.Shape.BlackWhiteMode
ms.assetid: 95a00870-82c2-d193-6971-9f92aeed6532
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.BlackWhiteMode property (Excel)

Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write **[MsoBlackWhiteMode](Office.MsoBlackWhiteMode.md)**.


## Syntax

_expression_.**BlackWhiteMode**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Example

This example sets shape one on `wksOne` to appear in black-and-white mode. When you view the presentation in black-and-white mode, shape one will appear black regardless of what color it is in color mode.

```vb
Sub UseBlackWhiteMode() 
 
    Dim wksOne As Worksheet 
    Set wksOne = Application.Worksheets(1) 
    wksOne.Shapes(1).BlackWhiteMode = msoBlackWhiteGrayOutline 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]