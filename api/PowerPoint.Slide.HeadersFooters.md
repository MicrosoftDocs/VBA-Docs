---
title: Slide.HeadersFooters property (PowerPoint)
keywords: vbapp10.chm531004
f1_keywords:
- vbapp10.chm531004
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.HeadersFooters
ms.assetid: 947eb2cf-6902-2eb1-f781-0602e96bbdef
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.HeadersFooters property (PowerPoint)

Returns a  **[HeadersFooters](PowerPoint.HeadersFooters.md)** collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides. Read-only.


## Syntax

_expression_. `HeadersFooters`

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Return value

HeadersFooters


## Example

This example sets the footer text and the date and time format for the notes master in the active presentation and sets the date and time to be updated automatically.


```vb
With ActivePresentation.NotesMaster.HeadersFooters

    .Footer.Text = "Regional Sales"

    With .DateAndTime

        .UseFormat = True

        .Format = ppDateTimeHmmss

    End With

End With
```


## See also


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]