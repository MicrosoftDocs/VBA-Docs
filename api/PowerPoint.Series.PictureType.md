---
title: Series.PictureType property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Series.PictureType
ms.assetid: 106933a2-49a7-e9d3-e5fa-fd2d0ab8974a
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.PictureType property (PowerPoint)

Returns or sets a value that specifies how pictures are displayed on a column or bar picture chart. Read/write  **[XlChartPictureType](PowerPoint.XlChartPictureType.md)**.


## Syntax

_expression_.**PictureType**

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Example

The following example sets series one of the first chart in the active document to stretch pictures. You should run the example on a 2D column chart that has picture data markers.




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).PictureType = xlStretch

    End If

End With
```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]