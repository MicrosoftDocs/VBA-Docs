---
title: Shapes.AddTitle method (PowerPoint)
keywords: vbapp10.chm543019
f1_keywords:
- vbapp10.chm543019
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddTitle
ms.assetid: 1fe13529-526a-1b29-7589-c155f9e46379
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddTitle method (PowerPoint)

Restores a previously deleted title placeholder to a slide. Returns a **[Shape](PowerPoint.Shape.md)** object that represents the restored title.


## Syntax

_expression_. `AddTitle`

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Return value

Shape


## Remarks

This method will cause an error if you haven't previously deleted the title placeholder from the specified slide. Use the  **[HasTitle](PowerPoint.Shapes.HasTitle.md)** property to determine whether the title placeholder has been deleted.


## Example

This example restores the title placeholder to slide one in the active presentation if this placeholder has been deleted. The text of the restored title is "Restored title."


```vb
With ActivePresentation.Slides(1)

    If .Layout <> ppLayoutBlank Then

        With .Shapes

            If Not .HasTitle Then

                .AddTitle.TextFrame.TextRange.Text = "Restored title"

            End If

        End With

    End If

End With
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]