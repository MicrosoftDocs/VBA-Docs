---
title: TextRange.Text property (PowerPoint)
keywords: vbapp10.chm569017
f1_keywords:
- vbapp10.chm569017
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Text
ms.assetid: c80c8b19-73e2-0820-abd6-c44f4b2644b2
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.Text property (PowerPoint)

Returns or sets a **String** that represents the text contained in the specified object. Read/write.


## Syntax

_expression_.**Text**

_expression_ A variable that represents a [TextRange](PowerPoint.TextRange.md) object.


## Return value

String


## Example

This example sets the text and font style for the title on slide one in the active presentation.


```vb
Set myPres = Application.ActivePresentation

With myPres.Slides(1).Shapes.Title.TextFrame.TextRange

    .Text = "Welcome!"

    .Font.Italic = True

End With
```


## See also


[TextRange Object](PowerPoint.TextRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]