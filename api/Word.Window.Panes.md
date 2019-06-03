---
title: Window.Panes property (Word)
keywords: vbawd10.chm157417475
f1_keywords:
- vbawd10.chm157417475
ms.prod: word
api_name:
- Word.Window.Panes
ms.assetid: d75cc2ab-940f-9e2b-81d5-bbbfdb0f4c6c
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.Panes property (Word)

Returns a  **[Panes](Word.panes.md)** collection that represents all the window panes for the specified window.


## Syntax

_expression_.**Panes**

 _expression_ An expression that returns a **[Window](Word.Window.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example splits the active window in half.


```vb
If ActiveDocument.ActiveWindow.Panes.Count = 1 Then _ 
 ActiveDocument.ActiveWindow.Panes.Add
```

This example activates the first pane in the window for Document2.




```vb
Windows("Document2").Panes(1).Activate
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]