---
title: Window.DisplayLeftScrollBar property (Word)
keywords: vbawd10.chm157417506
f1_keywords:
- vbawd10.chm157417506
ms.prod: word
api_name:
- Word.Window.DisplayLeftScrollBar
ms.assetid: 4f9be094-144c-cb4a-20e8-b3dc550a6bd0
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.DisplayLeftScrollBar property (Word)

 **True** if the vertical scroll bar appears on the left side of the document window. Read/write **Boolean**.


## Syntax

_expression_. `DisplayLeftScrollBar`

 _expression_ An expression that returns a **[Window](Word.Window.md)** object.


## Example

This example displays the vertical scroll bar on the left side of the active window.


```vb
ActiveWindow.DisplayLeftScrollBar = True
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]