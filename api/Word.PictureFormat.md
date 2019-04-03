---
title: PictureFormat object (Word)
keywords: vbawd10.chm2507
f1_keywords:
- vbawd10.chm2507
ms.prod: word
api_name:
- Word.PictureFormat
ms.assetid: 79556e36-81bb-f8df-45ef-c040df709497
ms.date: 06/08/2017
localization_priority: Normal
---


# PictureFormat object (Word)

Contains properties and methods that apply to pictures and OLE objects. The  **LinkFormat** object contains properties and methods that apply to linked OLE objects only. The **OLEFormat** object contains properties and methods that apply to OLE objects whether or not they're linked.


## Remarks

Use the  **PictureFormat** property to return a **PictureFormat** object. The following example sets the brightness, contrast, and color transformation for shape one on the active document and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.


```vb
With ActiveDocument.Shapes(1).PictureFormat 
 .Brightness = 0.3 
 .Contrast = 0.7 
 .ColorType = msoPictureGrayScale 
 .CropBottom = 18 
End With
```


## Methods



|Name|
|:-----|
|[IncrementBrightness](Word.PictureFormat.IncrementBrightness.md)|
|[IncrementContrast](Word.PictureFormat.IncrementContrast.md)|

## Properties



|Name|
|:-----|
|[Application](Word.PictureFormat.Application.md)|
|[Brightness](Word.PictureFormat.Brightness.md)|
|[ColorType](Word.PictureFormat.ColorType.md)|
|[Contrast](Word.PictureFormat.Contrast.md)|
|[Creator](Word.PictureFormat.Creator.md)|
|[Crop](Word.PictureFormat.Crop.md)|
|[CropBottom](Word.PictureFormat.CropBottom.md)|
|[CropLeft](Word.PictureFormat.CropLeft.md)|
|[CropRight](Word.PictureFormat.CropRight.md)|
|[CropTop](Word.PictureFormat.CropTop.md)|
|[Parent](Word.PictureFormat.Parent.md)|
|[TransparencyColor](Word.PictureFormat.TransparencyColor.md)|
|[TransparentBackground](Word.PictureFormat.TransparentBackground.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]