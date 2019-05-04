---
title: PageSetup.RightHeaderPicture property (Excel)
keywords: vbaxl10.chm473110
f1_keywords:
- vbaxl10.chm473110
ms.prod: excel
api_name:
- Excel.PageSetup.RightHeaderPicture
ms.assetid: 38fb53d1-7326-97d7-9c4a-285ffe8f42f7
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.RightHeaderPicture property (Excel)

Returns a **[Graphic](Excel.Graphic.md)** object that represents the picture for the right section of the header. Used to set attributes about the picture.

## Syntax

_expression_.**RightHeaderPicture**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

The **RightHeaderPicture** property is read-only, but not all of its properties are read-only.

It is required that `"&G"` be a part of the **RightHeader** property string for the image to show up in the right header.


## Example

The following example adds a picture titled Sample.jpg from the C:\ drive to the right section of the header. This example assumes that a file called Sample.jpg exists on the C:\ drive.

```vb
Sub InsertPicture() 
 
 With ActiveSheet.PageSetup.RightHeaderPicture 
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
 
 ' Enable the image to show up in the right header. 
 ActiveSheet.PageSetup.RightHeader = "&G" 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

