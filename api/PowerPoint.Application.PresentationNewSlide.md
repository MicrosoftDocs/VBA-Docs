---
title: Application.PresentationNewSlide event (PowerPoint)
keywords: vbapp10.chm621008
f1_keywords:
- vbapp10.chm621008
ms.prod: powerpoint
api_name:
- PowerPoint.Application.PresentationNewSlide
ms.assetid: e9718cad-6411-d013-6c93-0370aa71a8f2
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PresentationNewSlide event (PowerPoint)

Occurs when a new slide is created in any open presentation, as the slide is added to the  **[Slides](PowerPoint.Slides.md)** collection.


## Syntax

_expression_. `PresentationNewSlide`( `_Sld_` )

 _expression_ An expression that returns an **[Application](PowerPoint.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sld_|Required|**Slide**|The new slide.|

## Example

This example modifies the background color for color scheme three and then applies the modified color scheme to the new slide. Next, it adds default text to shape one if it has a text frame.


```vb
Private Sub App_PresentationNewSlide(ByVal Sld As Slide)

    With ActivePresentation

        Set CS3 = .ColorSchemes(3)

        CS3.Colors(ppBackground).RGB = RGB(240, 115, 100)

        Windows(1).Selection.SlideRange.ColorScheme = CS3

    End With



    If Sld.Layout <> ppLayoutBlank Then

        With Sld.Shapes(1)

            If .HasTextFrame = msoTrue Then

               .TextFrame.TextRange.Text = "King Salmon"

            End If

        End With

    End If

End Sub
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]