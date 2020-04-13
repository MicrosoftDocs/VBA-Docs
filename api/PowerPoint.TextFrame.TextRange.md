---
title: TextFrame.TextRange property (PowerPoint)
keywords: vbapp10.chm558008
f1_keywords:
- vbapp10.chm558008
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.TextRange
ms.assetid: 4a565e39-8bfa-7370-3ed6-57c442796144
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.TextRange property (PowerPoint)

Returns a **[TextRange](PowerPoint.TextRange.md)** object that represents the text in the specified text frame. Read-only.


## Syntax

_expression_.**TextRange**

_expression_ A variable that represents a **[TextFrame](PowerPoint.TextFrame.md)** object.


## Return value

TextRange


## Remarks

You can construct a text range from a selection when the presentation is in slide view, normal view, outline view, notes page view, or any master view.


## Example

This example makes the selected text bold in the first window.


```vb
Windows(1).Selection.TextRange.Font.Bold = True
```


## See also


[TextFrame Object](PowerPoint.TextFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]