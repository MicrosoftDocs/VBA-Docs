---
title: Pane.NewFrameset method (Word)
keywords: vbawd10.chm157286506
f1_keywords:
- vbawd10.chm157286506
ms.prod: word
api_name:
- Word.Pane.NewFrameset
ms.assetid: 86724851-6b29-1a66-e863-edeb4c9d43de
ms.date: 06/08/2017
localization_priority: Normal
---


# Pane.NewFrameset method (Word)

Creates a new frames page based on the specified pane.


## Syntax

_expression_. `NewFrameset`

_expression_ Required. A variable that represents a '[Pane](Word.Pane.md)' object.


## Remarks

For more information on creating frames pages, see [Creating frames pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## Example

This example opens a document named "Temp.doc" and then creates a new frames page whose only frame contains "Temp.doc".


```vb
Documents.Open "C:\Documents\Temp.doc" 
ActiveDocument.ActiveWindow.ActivePane.NewFrameset
```


## See also


[Pane Object](Word.Pane.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]