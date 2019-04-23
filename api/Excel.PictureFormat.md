---
title: PictureFormat object (Excel)
keywords: vbaxl10.chm113000
f1_keywords:
- vbaxl10.chm113000
ms.prod: excel
api_name:
- Excel.PictureFormat
ms.assetid: 7e8ec723-b6e0-fdc9-ff4e-22cbb31be4df
ms.date: 03/30/2019
localization_priority: Normal
---


# PictureFormat object (Excel)

Contains properties and methods that apply to pictures and OLE objects.


## Remarks

The **[LinkFormat](Excel.LinkFormat.md)** object contains properties and methods that apply to linked OLE objects only. The **[OLEFormat](Excel.OLEFormat.md)** object contains properties and methods that apply to OLE objects whether or not they're linked.


## Example

Use the **[PictureFormat](excel.shape.pictureformat.md)** property of the **Shape** object to return a **PictureFormat** object. The following example sets the brightness, contrast, and color transformation for shape one on _myDocument_ and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).PictureFormat 
 .Brightness = 0.3 
 .Contrast = 0.7 
 .ColorType = msoPictureGrayScale 
 .CropBottom = 18
```


## Methods

- [IncrementBrightness](Excel.PictureFormat.IncrementBrightness.md)
- [IncrementContrast](Excel.PictureFormat.IncrementContrast.md)

## Properties

- [Application](Excel.PictureFormat.Application.md)
- [Brightness](Excel.PictureFormat.Brightness.md)
- [ColorType](Excel.PictureFormat.ColorType.md)
- [Contrast](Excel.PictureFormat.Contrast.md)
- [Creator](Excel.PictureFormat.Creator.md)
- [Crop](Excel.PictureFormat.Crop.md)
- [CropBottom](Excel.PictureFormat.CropBottom.md)
- [CropLeft](Excel.PictureFormat.CropLeft.md)
- [CropRight](Excel.PictureFormat.CropRight.md)
- [CropTop](Excel.PictureFormat.CropTop.md)
- [Parent](Excel.PictureFormat.Parent.md)
- [TransparencyColor](Excel.PictureFormat.TransparencyColor.md)
- [TransparentBackground](Excel.PictureFormat.TransparentBackground.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]