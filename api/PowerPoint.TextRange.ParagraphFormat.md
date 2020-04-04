---
title: TextRange.ParagraphFormat property (PowerPoint)
keywords: vbapp10.chm569024
f1_keywords:
- vbapp10.chm569024
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.ParagraphFormat
ms.assetid: 41d3f0f3-70e3-ad1a-efcb-de849d4a03d4
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.ParagraphFormat property (PowerPoint)

Returns a **[ParagraphFormat](PowerPoint.ParagraphFormat.md)** object that represents paragraph formatting for the specified text. Read-only.


## Syntax

_expression_. `ParagraphFormat`

_expression_ A variable that represents a [TextRange](PowerPoint.TextRange.md) object.


## Return value

ParagraphFormat


## Example

This example sets the line spacing before, within, and after each paragraph in shape two on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(2).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat

        .LineRuleWithin = msoTrue

        .SpaceWithin = 1.4

        .LineRuleBefore = msoTrue

        .SpaceBefore = 0.25

        .LineRuleAfter = msoTrue

        .SpaceAfter = 0.75

    End With

End With
```


## See also


[TextRange Object](PowerPoint.TextRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]