---
title: Shape.PictureFormat property (Publisher)
keywords: vbapb10.chm2228295
f1_keywords:
- vbapb10.chm2228295
ms.prod: publisher
api_name:
- Publisher.Shape.PictureFormat
ms.assetid: 2a812ba3-18e4-fc42-6d07-535511a79650
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.PictureFormat property (Publisher)

Returns a **[PictureFormat](Publisher.PictureFormat.md)** object that contains picture formatting properties for the specified object. Applies to **Shape** or **[ShapeRange](Publisher.ShapeRange.md)** objects that represent pictures or OLE objects. Read-only.


## Syntax

_expression_.**PictureFormat**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Example

This example sets the brightness and contrast for all pictures on the first page of the active publication.

```vb
Sub FixPictureContrastBrightness() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 If shp.Type = pbPicture Then 
 With shp.PictureFormat 
 .Brightness = 0.6 
 .Contrast = 0.6 
 End With 
 End If 
 Next shp 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]