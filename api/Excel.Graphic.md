---
title: Graphic object (Excel)
keywords: vbaxl10.chm693072
f1_keywords:
- vbaxl10.chm693072
ms.prod: excel
api_name:
- Excel.Graphic
ms.assetid: 0ccdfb0d-effb-9fa4-8de9-b90688693375
ms.date: 06/08/2017
localization_priority: Normal
---


# Graphic object (Excel)

Contains properties that apply to header and footer picture objects.


## Remarks

There are several properties of the **[PageSetup](Excel.PageSetup.md)** object that return the **Graphic** object.

Use the **[CenterFooterPicture](Excel.PageSetup.CenterFooterPicture.md)**, **[CenterHeaderPicture](Excel.PageSetup.CenterHeaderPicture.md)**, **[LeftFooterPicture](Excel.PageSetup.LeftFooterPicture.md)**, **[LeftHeaderPicture](Excel.PageSetup.LeftHeaderPicture.md)**, **[RightFooterPicture](Excel.PageSetup.RightFooterPicture.md)**, or **[RightHeaderPicture](Excel.PageSetup.RightHeaderPicture.md)** properties to return a **Graphic** object.


 **Note**  It is required that "&G" is a part of the **LeftFooter** string in order for the image to show up in the left footer.


## Example

The following example adds a picture titled: Sample.jpg from the C:\ drive to the left section of the footer. This example assumes that a file called Sample.jpg exists on the C:\ drive.


```vb
Sub InsertPicture() 
 
 With ActiveSheet.PageSetup.LeftFooterPicture 
 .FileName = "C:\Sample.jpg" 
 .Height = 275.25 
 .Width = 463.5 
 .Brightness = 0.36 
 .ColorType = msoPictureGrayscale 
 .Contrast = 0.39 
 .CropBottom = -14.4 
 .CropLeft = -28.8 
 .CropRight = -14.4 
 .CropTop = 21.6 
 End With 
 
 ' Enable the image to show up in the left footer. 
 ActiveSheet.PageSetup.LeftFooter = "&G" 
 
End Sub
```

## Properties

- [Application](Excel.Graphic.Application.md)
- [Brightness](Excel.Graphic.Brightness.md)
- [ColorType](Excel.Graphic.ColorType.md)
- [Contrast](Excel.Graphic.Contrast.md)
- [Creator](Excel.Graphic.Creator.md)
- [CropBottom](Excel.Graphic.CropBottom.md)
- [CropLeft](Excel.Graphic.CropLeft.md)
- [CropRight](Excel.Graphic.CropRight.md)
- [CropTop](Excel.Graphic.CropTop.md)
- [Filename](Excel.Graphic.Filename.md)
- [Height](Excel.Graphic.Height.md)
- [LockAspectRatio](Excel.Graphic.LockAspectRatio.md)
- [Parent](Excel.Graphic.Parent.md)
- [Width](Excel.Graphic.Width.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]