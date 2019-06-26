---
title: Series.Paste method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Series.Paste
ms.assetid: 3f74aabb-f9c0-c76d-eaaa-c08c21daef48
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.Paste method (PowerPoint)

Pastes a picture from the Clipboard as the marker on the selected series.


## Syntax

_expression_.**Paste**

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Remarks

You can use this method on column, bar, line, or radar charts, and it sets the  **[MarkerStyle](PowerPoint.Series.MarkerStyle.md)** property to **xlMarkerStylePicture**.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example pastes a picture from the Clipboard into series one for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Paste

    End If

End With


```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]