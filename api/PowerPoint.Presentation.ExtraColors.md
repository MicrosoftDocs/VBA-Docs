---
title: Presentation.ExtraColors property (PowerPoint)
keywords: vbapp10.chm583014
f1_keywords:
- vbapp10.chm583014
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.ExtraColors
ms.assetid: c6a9d155-206c-36e6-c180-aaff8bd85a99
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.ExtraColors property (PowerPoint)

Returns an  **[ExtraColors](PowerPoint.ExtraColors.md)** object that represents the extra colors available in the specified presentation. Read-only.


## Syntax

_expression_. `ExtraColors`

_expression_ A variable that represents an [Presentation](PowerPoint.Presentation.md) object.


## Return value

ExtraColors


## Example

The following example adds a rectangle to slide one in the active presentation and sets its fill foreground color to the first extra color. If there hasn't been at least one extra color defined for the presentation, this example will fail.


```vb
With ActivePresentation
    Set rect = .Slides(1).Shapes _
        .AddShape(msoShapeRectangle, 50, 50, 100, 200)
    rect.Fill.ForeColor.RGB = .ExtraColors(1)
End With
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]