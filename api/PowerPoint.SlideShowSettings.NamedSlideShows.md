---
title: SlideShowSettings.NamedSlideShows property (PowerPoint)
keywords: vbapp10.chm514004
f1_keywords:
- vbapp10.chm514004
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.NamedSlideShows
ms.assetid: 8af7610f-1981-df5f-5be8-2bb04c895602
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowSettings.NamedSlideShows property (PowerPoint)

Returns a **[NamedSlideShows](PowerPoint.NamedSlideShows.md)** collection that represents all the named slide shows (custom slide shows) in the specified presentation. Read-only.


## Syntax

_expression_. `NamedSlideShows`

_expression_ A variable that represents a [SlideShowSettings](PowerPoint.SlideShowSettings.md) object.


## Return value

NamedSlideShows


## Remarks

Each named slide show, or custom slide show, is a user-selected subset of the specified presentation.

Use the  **[Add](PowerPoint.NamedSlideShows.Add.md)** method of the **NamedSlideShows** object to create a named slide show.


## Example

This example adds to the active presentation a named slide show "Quick Show" that contains slides 2, 7, and 9. The example then runs this slide show.


```vb
Dim qSlides(1 To 3) As Long

With ActivePresentation

    With .Slides

        qSlides(1) = .Item(2).SlideID

        qSlides(2) = .Item(7).SlideID

        qSlides(3) = .Item(9).SlideID

    End With

    With .SlideShowSettings

        .RangeType = ppShowNamedSlideShow

        .NamedSlideShows.Add "Quick Show", qSlides

        .SlideShowName = "Quick Show"

        .Run

    End With

End With
```


## See also


[SlideShowSettings Object](PowerPoint.SlideShowSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]