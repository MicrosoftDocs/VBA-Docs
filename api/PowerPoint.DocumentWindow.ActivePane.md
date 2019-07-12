---
title: DocumentWindow.ActivePane property (PowerPoint)
keywords: vbapp10.chm511022
f1_keywords:
- vbapp10.chm511022
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.ActivePane
ms.assetid: 8fa4c8a1-37b6-2676-1cfd-5fa2b130d2e3
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentWindow.ActivePane property (PowerPoint)

Returns a  **[Pane](PowerPoint.Pane.md)** object that represents the active pane in the document window. Read-only.


## Syntax

_expression_.**ActivePane**

_expression_ A variable that represents an [DocumentWindow](PowerPoint.DocumentWindow.md) object.


## Return value

Pane


## Example

If the active pane is the slide pane, this example makes the notes pane the active pane. The notes pane is the third member of the  **Panes** collection.


```vb
With ActiveWindow

    If .ActivePane.ViewType = ppViewSlide Then

        .Panes(3).Activate

    End If

End With
```


## See also



[DocumentWindow Object](PowerPoint.DocumentWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]