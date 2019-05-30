---
title: Window.View property (Word)
keywords: vbawd10.chm157417486
f1_keywords:
- vbawd10.chm157417486
ms.prod: word
api_name:
- Word.Window.View
ms.assetid: d012af14-e1cc-b13e-e1d1-48ea53ba0f0a
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.View property (Word)

Returns a  **[View](Word.View.md)** object that represents the view for the specified window or pane.


## Syntax

_expression_.**View**

_expression_ Required. A variable that represents a **[Window](Word.Window.md)** object.


## Example

This example switches the active window to full-screen view.


```vb
ActiveDocument.ActiveWindow.View.FullScreen = True
```

This example sets view options for each window in the  **Windows** collection.




```vb
For Each myWindow In Windows 
 With myWindow.View 
 .ShowTabs = True 
 .ShowParagraphs = True 
 .Type = wdNormalView 
 End With 
Next myWindow
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]