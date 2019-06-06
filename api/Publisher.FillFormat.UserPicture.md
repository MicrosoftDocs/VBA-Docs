---
title: FillFormat.UserPicture method (Publisher)
keywords: vbapb10.chm2359319
f1_keywords:
- vbapb10.chm2359319
ms.prod: publisher
api_name:
- Publisher.FillFormat.UserPicture
ms.assetid: b1eaf724-42b4-657f-4d88-bc8547664893
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.UserPicture method (Publisher)

Fills the specified shape with one large image.


## Syntax

_expression_.**UserPicture** (_PictureFile_)

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PictureFile_|Required| **String**|The name of the picture file.|

## Remarks

To fill the shape with small tiles of an image, use the **[UserTextured](Publisher.FillFormat.UserTextured.md)** method.


## Example

This example adds two rectangles to the active publication. The rectangle on the left is filled with one large image of a picture; the rectangle on the right is filled with many small tiles of the same picture. Note that `PathToFile` must be replaced with a valid file path for this example to work.

```vb
With ActiveDocument.Pages(1).Shapes 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=200, Height:=100).Fill _ 
 .UserPicture PictureFile:="PathToFile" 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=300, Top:=0, Width:=200, Height:=100).Fill _ 
 .UserTextured TextureFile:="PathToFile" 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]