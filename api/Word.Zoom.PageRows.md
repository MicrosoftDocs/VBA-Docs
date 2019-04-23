---
title: Zoom.PageRows property (Word)
keywords: vbawd10.chm161873922
f1_keywords:
- vbawd10.chm161873922
ms.prod: word
api_name:
- Word.Zoom.PageRows
ms.assetid: 15db7c14-ee98-bac7-179a-018f4cb47fb9
ms.date: 06/08/2017
localization_priority: Normal
---


# Zoom.PageRows property (Word)

Returns or sets the number of pages to be displayed one above the other on-screen at the same time in print layout view or print preview. Read/write  **Long**.


## Syntax

_expression_. `PageRows`

 _expression_ An expression that returns a '[Zoom](Word.Zoom.md)' object.


## Example

This example switches the active window to print preview and displays two pages one above the other.


```vb
PrintPreview = True 
With ActiveDocument.ActiveWindow.View.Zoom 
 .PageColumns = 1 
 .PageRows = 2 
End With
```


## See also


[Zoom Object](Word.Zoom.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]