---
title: ShapeRange.PictureFormat property (Publisher)
keywords: vbapb10.chm2293831
f1_keywords:
- vbapb10.chm2293831
ms.prod: publisher
api_name:
- Publisher.ShapeRange.PictureFormat
ms.assetid: 3d693c6b-b76b-0fe1-e7df-63fb08782f6f
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.PictureFormat property (Publisher)

Returns a **[PictureFormat](Publisher.PictureFormat.md)** object that contains picture formatting properties for the specified object. Applies to **[Shape](Publisher.Shape.md)** or **ShapeRange** objects that represent pictures or OLE objects. Read-only.


## Syntax

_expression_.**PictureFormat**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


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