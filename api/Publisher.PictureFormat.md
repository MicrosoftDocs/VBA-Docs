---
title: PictureFormat object (Publisher)
keywords: vbapb10.chm3670015
f1_keywords:
- vbapb10.chm3670015
ms.prod: publisher
api_name:
- Publisher.PictureFormat
ms.assetid: aa30ea9d-b91f-acdf-2e60-8a9f506f28b4
ms.date: 06/01/2019
localization_priority: Normal
---


# PictureFormat object (Publisher)

Contains properties and methods that apply to pictures.

## Remarks

Use the **[Shape.PictureFormat](Publisher.Shape.PictureFormat.md)** property to return a **PictureFormat** object. 

## Example

The following example sets the brightness, contrast, and color transformation for shape one on the active document and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.

```vb
Sub FormatPicture() 
 With ActiveDocument.Pages(1).Shapes(1).PictureFormat 
 .Brightness = 0.6 
 .Contrast = 0.7 
 .ColorType = msoPictureGrayscale 
 .CropBottom = 18 
 End With 
End Sub
```


## Methods

- [ClearCrop](Publisher.PictureFormat.ClearCrop.md)
- [FillFrame](Publisher.PictureFormat.FillFrame.md)
- [FitFrame](Publisher.PictureFormat.FitFrame.md)
- [IncrementBrightness](Publisher.PictureFormat.IncrementBrightness.md)
- [IncrementContrast](Publisher.PictureFormat.IncrementContrast.md)
- [Recolor](Publisher.PictureFormat.Recolor.md)
- [Remove](Publisher.PictureFormat.Remove.md)
- [Replace](Publisher.PictureFormat.Replace.md)
- [ReplaceEx](Publisher.PictureFormat.ReplaceEx.md)
- [RestoreOriginalColors](Publisher.PictureFormat.RestoreOriginalColors.md)

## Properties

- [Application](Publisher.PictureFormat.Application.md)
- [Brightness](Publisher.PictureFormat.Brightness.md)
- [ColorModel](Publisher.PictureFormat.ColorModel.md)
- [ColorsInPalette](Publisher.PictureFormat.ColorsInPalette.md)
- [ColorType](Publisher.PictureFormat.ColorType.md)
- [Contrast](Publisher.PictureFormat.Contrast.md)
- [CropBottom](Publisher.PictureFormat.CropBottom.md)
- [CropLeft](Publisher.PictureFormat.CropLeft.md)
- [CropRight](Publisher.PictureFormat.CropRight.md)
- [CropTop](Publisher.PictureFormat.CropTop.md)
- [EffectiveResolution](Publisher.PictureFormat.EffectiveResolution.md)
- [FileName](Publisher.PictureFormat.FileName.md)
- [FileSize](Publisher.PictureFormat.FileSize.md)
- [HasAlphaChannel](Publisher.PictureFormat.HasAlphaChannel.md)
- [HasTransparencyColor](Publisher.PictureFormat.HasTransparencyColor.md)
- [Height](Publisher.PictureFormat.Height.md)
- [HorizontalPictureLocking](Publisher.PictureFormat.HorizontalPictureLocking.md)
- [HorizontalScale](Publisher.PictureFormat.HorizontalScale.md)
- [ImageFormat](Publisher.PictureFormat.ImageFormat.md)
- [IsEmpty](Publisher.PictureFormat.IsEmpty.md)
- [IsGreyScale](Publisher.PictureFormat.IsGreyScale.md)
- [IsLinked](Publisher.PictureFormat.IsLinked.md)
- [IsRecolored](Publisher.PictureFormat.IsRecolored.md)
- [IsTrueColor](Publisher.PictureFormat.IsTrueColor.md)
- [LeaveBlackAsBlack](Publisher.PictureFormat.LeaveBlackAsBlack.md)
- [LinkedFileStatus](Publisher.PictureFormat.LinkedFileStatus.md)
- [OriginalColorsInPalette](Publisher.PictureFormat.OriginalColorsInPalette.md)
- [OriginalFileSize](Publisher.PictureFormat.OriginalFileSize.md)
- [OriginalHasAlphaChannel](Publisher.PictureFormat.OriginalHasAlphaChannel.md)
- [OriginalHeight](Publisher.PictureFormat.OriginalHeight.md)
- [OriginalIsTrueColor](Publisher.PictureFormat.OriginalIsTrueColor.md)
- [OriginalResolution](Publisher.PictureFormat.OriginalResolution.md)
- [OriginalWidth](Publisher.PictureFormat.OriginalWidth.md)
- [Parent](Publisher.PictureFormat.Parent.md)
- [RecoloredPictureColor](Publisher.PictureFormat.RecoloredPictureColor.md)
- [TransparencyColor](Publisher.PictureFormat.TransparencyColor.md)
- [TransparentBackground](Publisher.PictureFormat.TransparentBackground.md)
- [VerticalPictureLocking](Publisher.PictureFormat.VerticalPictureLocking.md)
- [VerticalScale](Publisher.PictureFormat.VerticalScale.md)
- [Width](Publisher.PictureFormat.Width.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]