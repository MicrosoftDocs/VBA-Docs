---
title: Series.PictureUnit2 property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Series.PictureUnit2
ms.assetid: 83ccb10a-1883-9665-8a63-4494e853aa72
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.PictureUnit2 property (PowerPoint)

Returns or sets the unit for each picture on the chart if the  **[PictureType](PowerPoint.Series.PictureType.md)** property is set to **xlStackScale**; otherwise, this property is ignored. Read/write **Double**.


## Syntax

_expression_.**PictureUnit2**

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets series one for the first chart in the active document to stack pictures and uses each picture to represent five units. You should run the example on a 2D column chart that has picture data markers.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1)

            .PictureType = xlScale

            .PictureUnit2 = 5

        End With

    End If

End With
```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]