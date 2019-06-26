---
title: Series.MarkerSize property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Series.MarkerSize
ms.assetid: 60a402b8-69f5-db47-73df-55ed75a42272
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.MarkerSize property (PowerPoint)

Returns or sets the data-marker size, in points. Read/write  **Long**.


## Syntax

_expression_.**MarkerSize**

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Remarks

This property can have a value from 2 through 72. 


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the data-marker size for all data markers on series one for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).MarkerSize = 10

    End If

End With


```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]