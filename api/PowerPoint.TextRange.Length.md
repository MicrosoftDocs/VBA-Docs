---
title: TextRange.Length property (PowerPoint)
keywords: vbapp10.chm569005
f1_keywords:
- vbapp10.chm569005
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Length
ms.assetid: 4eb64830-f8e4-5226-57c1-80df7f4bd39f
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.Length property (PowerPoint)

Returns the length of the specified text range, in characters. Read-only.


## Syntax

_expression_.**Length**

_expression_ A variable that represents a [TextRange](PowerPoint.TextRange.md) object.


## Return value

Long


## Example

This example sets the title font size to 48 points if the title of slide two contains more than five characters, or it sets the font size to 72 points if the title has five or fewer characters.


```vb
Set myDocument = ActivePresentation.Slides(2)

With myDocument.Shapes(1).TextFrame.TextRange

    If .Length > 5 Then

        .Font.Size = 48

    Else

        .Font.Size = 72

    End If

End With


```


## See also


[TextRange Object](PowerPoint.TextRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]