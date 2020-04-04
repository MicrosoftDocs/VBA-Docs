---
title: TextRange.Font property (PowerPoint)
keywords: vbapp10.chm569023
f1_keywords:
- vbapp10.chm569023
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Font
ms.assetid: 234c8843-3c0d-a425-0173-02c3910ba400
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.Font property (PowerPoint)

Returns a **[Font](PowerPoint.Font.md)** object that represents character formatting. Read-only.


## Syntax

_expression_.**Font**

_expression_ A variable that represents a [TextRange](PowerPoint.TextRange.md) object.


## Return value

Font


## Example

This example sets the formatting for the text in shape one on slide one in the active presentation.


```vb
With ActivePresentation.Slides(1).Shapes(1)

    With .TextFrame.TextRange.Font

        .Size = 48

        .Name = "Palatino"

        .Bold = True

        .Color.RGB = RGB(255, 127, 255)

    End With

End With
```

This example sets the color and font name for bullets in shape two on slide one.




```vb
With ActivePresentation.Slides(1).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat.Bullet

        .Visible = True

        With .Font

            .Name = "Palatino"

            .Color.RGB = RGB(0, 0, 255)

        End With

    End With

End With
```


## See also


[TextRange Object](PowerPoint.TextRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
