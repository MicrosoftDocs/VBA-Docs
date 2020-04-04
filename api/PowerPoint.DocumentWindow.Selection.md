---
title: DocumentWindow.Selection property (PowerPoint)
keywords: vbapp10.chm511003
f1_keywords:
- vbapp10.chm511003
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.Selection
ms.assetid: 0cd670b2-53a5-87d7-8b38-761920dd9758
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentWindow.Selection property (PowerPoint)

Returns a **[Selection](PowerPoint.Selection.md)** object that represents the selection in the specified document window. Read-only.


## Syntax

_expression_.**Selection**

_expression_ A variable that represents a [DocumentWindow](PowerPoint.DocumentWindow.md) object.


## Return value

Selection


## Example

If there's text selected in the active window, this example makes the text italic.


```vb
With Application.ActiveWindow.Selection

    If .Type = ppSelectionText Then

        .TextRange.Font.Italic = True

    End If

End With


```


## See also



[DocumentWindow Object](PowerPoint.DocumentWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]