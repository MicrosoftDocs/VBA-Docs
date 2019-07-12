---
title: DocumentWindow.Panes property (PowerPoint)
keywords: vbapp10.chm511023
f1_keywords:
- vbapp10.chm511023
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.Panes
ms.assetid: 1f26709d-8414-ee89-29d8-588c6787611a
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentWindow.Panes property (PowerPoint)

Returns a  **[Panes](PowerPoint.Panes.md)** collection that represents the panes in the document window. Read-only.


## Syntax

_expression_.**Panes**

_expression_ A variable that represents a [DocumentWindow](PowerPoint.DocumentWindow.md) object.


## Return value

Panes


## Example

This example tests for the number of panes in the active window. If the value is one, indicating any view other that normal view, normal view is activated.


```vb
If ActiveWindow.Panes.Count = 1 Then

    ActiveWindow.ViewType = ppViewNormal

End If
```


## See also


[DocumentWindow Object](PowerPoint.DocumentWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]