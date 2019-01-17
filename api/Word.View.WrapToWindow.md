---
title: View.WrapToWindow property (Word)
keywords: vbawd10.chm161808404
f1_keywords:
- vbawd10.chm161808404
ms.prod: word
api_name:
- Word.View.WrapToWindow
ms.assetid: f596f4e6-c404-3b58-93a8-8aca79b60b66
ms.date: 06/08/2017
localization_priority: Normal
---


# View.WrapToWindow property (Word)

 **True** if lines wrap at the right edge of the document window rather than at the right margin or the right column boundary. Read/write **Boolean**.


## Syntax

 _expression_. `WrapToWindow`

 _expression_ An expression that returns a '[View](Word.View.md)' object.


## Remarks

This property has no effect in print layout or Web layout view.


## Example

This example wraps the text to fit within the active window.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdNormalView 
 .WrapToWindow = True 
End With
```


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]