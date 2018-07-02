---
title: Crop Object (Office)
ms.prod: office
api_name:
- Office.Crop
ms.assetid: 21ac150e-0a8f-c77b-717f-bf38fbced5a3
ms.date: 06/08/2017
---


# Crop Object (Office)

An object used to remove a portion of an image.


## Example

The following example inserts a 200 x 200 image into a PowerPoint presentation approximately in the center of the slide. It then resizes the image inside the frame to 100 x 100. The image frame stays at 200 x 200. The code then adds a square (the default shape) just above and to the right of the image, essentially cropping the lower left corner of the image.


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


## Properties



|**Name**|
|:-----|
|[Application](Office.Crop.Application.md)|
|[Creator](Office.Crop.Creator.md)|
|[PictureHeight](Office.Crop.PictureHeight.md)|
|[PictureOffsetX](Office.Crop.PictureOffsetX.md)|
|[PictureOffsetY](Office.Crop.PictureOffsetY.md)|
|[PictureWidth](Office.Crop.PictureWidth.md)|
|[ShapeHeight](Office.Crop.ShapeHeight.md)|
|[ShapeLeft](Office.Crop.ShapeLeft.md)|
|[ShapeTop](Office.Crop.ShapeTop.md)|
|[ShapeWidth](Office.Crop.ShapeWidth.md)|

## See also





[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
