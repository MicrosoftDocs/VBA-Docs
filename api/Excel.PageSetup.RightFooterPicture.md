---
title: PageSetup.RightFooterPicture property (Excel)
keywords: vbaxl10.chm473111
f1_keywords:
- vbaxl10.chm473111
ms.prod: excel
api_name:
- Excel.PageSetup.RightFooterPicture
ms.assetid: f33bbfb1-91d0-6724-0944-2b63c6720d86
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.RightFooterPicture property (Excel)

Returns a **[Graphic](Excel.Graphic.md)** object that represents the picture for the right section of the footer. Used to set attributes of the picture.


## Syntax

_expression_.**RightFooterPicture**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

The **RightFooterPicture** property itself is read-only, but not all of its properties are read-only.

It is required that `"&G"` be a part of the **RightFooter** property string for the image to show up in the right footer.


## Example

The following example adds a picture titled Sample.jpg from the C:\ drive to the right section of the footer. This example assumes that a file called Sample.jpg exists on the C:\ drive.

```vb
Sub InsertPicture() 
 
 With ActiveSheet.PageSetup.RightFooterPicture 
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
 
 ' Enable the image to show up in the right footer. 
 ActiveSheet.PageSetup.RightFooter = "&G" 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]