---
title: ParagraphFormat.SpaceAfter Property (PowerPoint)
keywords: vbapp10.chm576009
f1_keywords:
- vbapp10.chm576009
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.SpaceAfter
ms.assetid: 8b5dcf96-c35f-5e0b-6bd2-efabce7ea16f
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.SpaceAfter Property (PowerPoint)

Returns or sets the amount of space after the last line in each paragraph of the specified text, in points or lines. Read/write.


## Syntax

 _expression_. `SpaceAfter`

 _expression_ A variable that represents a [ParagraphFormat](./PowerPoint.ParagraphFormat.md) object.


## Return value

Single


## Example

This example sets the spacing after paragraphs to 6 points for the text in shape two on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat

        .LineRuleAfter = False

        .SpaceAfter = 6

    End With

End With
```


## See also


[ParagraphFormat Object](PowerPoint.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]