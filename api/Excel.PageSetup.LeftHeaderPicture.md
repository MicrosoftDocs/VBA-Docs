---
title: PageSetup.LeftHeaderPicture property (Excel)
keywords: vbaxl10.chm473108
f1_keywords:
- vbaxl10.chm473108
ms.prod: excel
api_name:
- Excel.PageSetup.LeftHeaderPicture
ms.assetid: 1dadb662-c93c-5fdb-ffef-24978284d35a
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.LeftHeaderPicture property (Excel)

Returns a  **[Graphic](Excel.Graphic.md)** object that represents the picture for the left section of the header. Used to set attributes about the picture.


## Syntax

_expression_. `LeftHeaderPicture`

_expression_ A variable that represents a [PageSetup](Excel.PageSetup.md) object.


## Remarks

The  **LeftHeaderPicture** property is read-only, but not all of its properties are read-only.


 **Note**  It is required that "&G" be a part of the  **LeftHeader** property string in order for the image to show up in the left header.


## Example

The following example adds a picture titled: Sample.jpg from the C: drive to the left section of the header. This example assumes that a file called Sample.jpg exists on the C: drive.


```vb
Sub InsertPicture() 
 
 With ActiveSheet.PageSetup.LeftHeaderPicture 
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
 
 ' Enable the image to show up in the left header. 
 ActiveSheet.PageSetup.LeftHeader = "&G" 
 
End Sub
```


## See also


[PageSetup Object](Excel.PageSetup.md)

