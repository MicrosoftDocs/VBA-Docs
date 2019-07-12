---
title: PictureFormat object (PowerPoint)
keywords: vbapp10.chm551000
f1_keywords:
- vbapp10.chm551000
ms.prod: powerpoint
api_name:
- PowerPoint.PictureFormat
ms.assetid: 946794b4-0401-ec7c-cea3-779ebfce0d69
ms.date: 06/08/2017
localization_priority: Normal
---


# PictureFormat object (PowerPoint)

Contains properties and methods that apply to pictures and OLE objects. 


## Example

Use the  **PictureFormat** property to return a **PictureFormat** object. The following example sets the brightness, contrast, and color transformation for shape one on _myDocument_ and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).PictureFormat

    .Brightness = 0.3

    .Contrast = 0.7

    .ColorType = msoPictureGrayScale

    .CropBottom = 18

End With
```


## Methods



|Name|
|:-----|
|[IncrementBrightness](PowerPoint.PictureFormat.IncrementBrightness.md)|
|[IncrementContrast](PowerPoint.PictureFormat.IncrementContrast.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.PictureFormat.Application.md)|
|[Brightness](PowerPoint.PictureFormat.Brightness.md)|
|[ColorType](PowerPoint.PictureFormat.ColorType.md)|
|[Contrast](PowerPoint.PictureFormat.Contrast.md)|
|[Creator](PowerPoint.PictureFormat.Creator.md)|
|[Crop](PowerPoint.PictureFormat.Crop.md)|
|[CropBottom](PowerPoint.PictureFormat.CropBottom.md)|
|[CropLeft](PowerPoint.PictureFormat.CropLeft.md)|
|[CropRight](PowerPoint.PictureFormat.CropRight.md)|
|[CropTop](PowerPoint.PictureFormat.CropTop.md)|
|[Parent](PowerPoint.PictureFormat.Parent.md)|
|[TransparencyColor](PowerPoint.PictureFormat.TransparencyColor.md)|
|[TransparentBackground](PowerPoint.PictureFormat.TransparentBackground.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]