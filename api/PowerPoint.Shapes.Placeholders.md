---
title: Shapes.Placeholders property (PowerPoint)
keywords: vbapp10.chm543021
f1_keywords:
- vbapp10.chm543021
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.Placeholders
ms.assetid: 2926d893-056a-0805-85ba-681e64bf81ed
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.Placeholders property (PowerPoint)

Returns a **[Placeholders](PowerPoint.Placeholders.md)** collection that represents the collection of all the placeholders on a slide. Read-only.


## Syntax

_expression_. `Placeholders`

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Return value

Placeholders


## Remarks

Each placeholder in the  **Placeholders** collection can contain text, a chart, a table, an organizational chart, or another object.


## Example

This example adds a slide to the active presentation and then adds text to both the title (which is the first placeholder on the slide) and the subtitle.


```vb
Set myDocument = ActivePresentation.Slides(1)

With ActivePresentation.Slides _
        .Add(1, ppLayoutTitle).Shapes.Placeholders

    .Item(1).TextFrame.TextRange.Text = "This is the title text"
    .Item(2).TextFrame.TextRange.Text = "This is subtitle text"

End With
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]