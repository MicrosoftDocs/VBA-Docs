---
title: Window.Thumbnails property (Word)
keywords: vbawd10.chm157417509
f1_keywords:
- vbawd10.chm157417509
ms.prod: word
api_name:
- Word.Window.Thumbnails
ms.assetid: 2979b109-e2e6-34de-539b-53c46b0d0c55
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.Thumbnails property (Word)

Sets or returns a  **Boolean** that represents whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.


## Syntax

_expression_. `Thumbnails`

 _expression_ An expression that returns a **[Window](Word.Window.md)** object.


## Example

The following example displays thumbnail images of the pages in the active document.


```vb
ActiveDocument.ActiveWindow.Thumbnails = True
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]