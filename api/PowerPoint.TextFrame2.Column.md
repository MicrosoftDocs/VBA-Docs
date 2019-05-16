---
title: TextFrame2.Column property (PowerPoint)
keywords: vbapp10.chm678017
f1_keywords:
- vbapp10.chm678017
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.Column
ms.assetid: d265fd2c-1e96-984d-9b2c-0a792cbf7671
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.Column property (PowerPoint)

Returns the  **[Column](PowerPoint.Column.md)** object that represents the columns of the specified text frame. Read-only.


## Syntax

_expression_.**Column**

 _expression_ An expression that returns a **[TextFrame2](PowerPoint.TextFrame2.md)** object.


## Example

The following example shows how to set the number of columns in the text frame of the first shape on slide one to 2.


```vb
Public Sub Column_Example()

    ActivePresentation.Slides(1).Shapes(1).TextFrame2.Column.Number = 2

End Sub
```


## See also


[TextFrame2 Object](PowerPoint.TextFrame2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]