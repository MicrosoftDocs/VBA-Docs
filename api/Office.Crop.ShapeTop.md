---
title: Crop.ShapeTop property (Office)
ms.prod: office
api_name:
- Office.Crop.ShapeTop
ms.assetid: 4bbb6765-e571-0d66-d6e6-96cc54e67d3f
ms.date: 01/04/2019
localization_priority: Normal
---


# Crop.ShapeTop property (Office)

Gets or sets the location of the top of a shape that is used to crop an image. Read/write.


## Syntax

_expression_.**ShapeTop**

_expression_ An expression that returns a **[Crop](Office.Crop.md)** object.


## Return value

Single


## Remarks

In Microsoft Word, the **ShapeLeft** and **ShapeTop** properties will return an error if the picture or shape has the wrap text set to **Inline with Text**.


## Example

The following example inserts a 200 x 200 image into a PowerPoint presentation approximately in the center of the slide. It then resizes the image inside the frame to 100 x 100. The image frame stays at 200 x 200. The code then adds a square (the default shape) just above and to the right of the image, essentially cropping the lower-left corner of the image.


```vb
Sub CropImage() 
 ActivePresentation.Slides(1).Shapes.AddPicture "c:\myImage.png", msoFalse, msoTrue, 250,150, 200, 200 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureHeight = 100 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureWidth = 100 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureOffsetX = 0 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.PictureOffsetY = 0 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeHeight = 100 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeWidth = 100 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeLeft = 330 
 ActivePresentation.Slides(1).Shapes(1).PictureFormat.Crop.ShapeTop = 170 
End Sub 

```


## See also

- [Crop object members](overview/library-reference/crop-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]