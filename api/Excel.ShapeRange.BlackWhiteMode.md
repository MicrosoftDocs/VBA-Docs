---
title: ShapeRange.BlackWhiteMode property (Excel)
keywords: vbaxl10.chm640124
f1_keywords:
- vbaxl10.chm640124
ms.prod: excel
api_name:
- Excel.ShapeRange.BlackWhiteMode
ms.assetid: df88d789-6686-2f02-1e69-54c8ab47060c
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.BlackWhiteMode property (Excel)

Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write **[MsoBlackWhiteMode](Office.MsoBlackWhiteMode.md)**.


## Syntax

_expression_.**BlackWhiteMode**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


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