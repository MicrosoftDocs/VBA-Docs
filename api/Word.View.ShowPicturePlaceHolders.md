---
title: View.ShowPicturePlaceHolders property (Word)
keywords: vbawd10.chm161808405
f1_keywords:
- vbawd10.chm161808405
ms.prod: word
api_name:
- Word.View.ShowPicturePlaceHolders
ms.assetid: 6a3d1529-57ab-eb56-225e-dee87ebc1185
ms.date: 06/08/2017
localization_priority: Normal
---


# View.ShowPicturePlaceHolders property (Word)

 **True** if blank boxes are displayed as placeholders for pictures. Read/write **Boolean**.


## Syntax

_expression_. `ShowPicturePlaceHolders`

 _expression_ An expression that returns a '[View](Word.View.md)' object.


## Example

This example inserts a picture in the active document and displays picture placeholders in the active window.


```vb
Selection.Collapse Direction:=wdCollapseStart 
ActiveDocument.InlineShapes.AddPicture Range:=Selection.Range, _ 
 FileName:="C:\Windows\Bubbles.bmp" 
ActiveDocument.ActiveWindow.View.ShowPicturePlaceHolders = True
```


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]