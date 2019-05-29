---
title: Window.SplitVertical property (Word)
keywords: vbawd10.chm157417482
f1_keywords:
- vbawd10.chm157417482
ms.prod: word
api_name:
- Word.Window.SplitVertical
ms.assetid: db04a1d5-0f5a-d17c-6a47-1da6b0e7f124
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.SplitVertical property (Word)

Returns or sets the vertical split percentage for the specified window. Read/write  **Long**.


## Syntax

_expression_.**SplitVertical**

 _expression_ An expression that returns a **[Window](Word.Window.md)** object.


## Remarks

To remove the split, set this property to zero (0) or set the  **[Split](Word.Window.Split.md)** property to **False**.


## Example

This example splits the active window so that the top pane occupies 70 percent of the window.


```vb
ActiveDocument.ActiveWindow.SplitVertical = 70
```

This example splits the window for Document1 in half vertically.




```vb
Windows("Document1").SplitVertical = 50
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]