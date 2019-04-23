---
title: View.ShowObjectAnchors property (Word)
keywords: vbawd10.chm161808395
f1_keywords:
- vbawd10.chm161808395
ms.prod: word
api_name:
- Word.View.ShowObjectAnchors
ms.assetid: 6b3c0f7a-0bf2-8671-1281-6ef61ae62ef8
ms.date: 06/08/2017
localization_priority: Normal
---


# View.ShowObjectAnchors property (Word)

 **True** if object anchors are displayed next to items that can be positioned in print layout view. Read/write **Boolean**.


## Syntax

_expression_. `ShowObjectAnchors`

 _expression_ An expression that returns a '[View](Word.View.md)' object.


## Example

This example adds a frame around the selection, switches the active window to print layout view, and shows object anchors for framed objects.


```vb
Selection.Frames.Add(Range:=Selection.Range).LockAnchor = True 
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .ShowObjectAnchors = True 
End With
```


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]