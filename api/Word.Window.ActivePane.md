---
title: Window.ActivePane property (Word)
keywords: vbawd10.chm157417473
f1_keywords:
- vbawd10.chm157417473
ms.prod: word
api_name:
- Word.Window.ActivePane
ms.assetid: 8491d406-5444-2d11-da29-8de575a0e066
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.ActivePane property (Word)

Returns a  **[Pane](Word.Pane.md)** object that represents the active pane for the specified window. Read-only.


## Syntax

 _expression_. `ActivePane`

 _expression_ A variable that represents a '[Window](Word.Window.md)' object.


## Example

This example splits the active window and then activates the next pane after the active pane.


```vb
With ActiveDocument.ActiveWindow 
 .Split = True 
 .ActivePane.Next.Activate 
 MsgBox "Pane " & .ActivePane.Index & " is active" 
End With
```

This example activates the first window and displays tabs in the active pane.




```vb
With Application.Windows(1) 
 .Activate 
 .ActivePane.View.ShowTabs = True 
End With
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]