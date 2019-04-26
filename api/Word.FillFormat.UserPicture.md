---
title: FillFormat.UserPicture method (Word)
keywords: vbawd10.chm164102161
f1_keywords:
- vbawd10.chm164102161
ms.prod: word
api_name:
- Word.FillFormat.UserPicture
ms.assetid: 09ddb55f-7ba0-9345-c366-23ac5ce6945a
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.UserPicture method (Word)

Fills the specified shape with one large image. 


## Syntax

_expression_.**UserPicture** (_PictureFile_)

_expression_ Required. A variable that represents a **[FillFormat](word.fillformat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PictureFile_|Required| **String**|The name of the picture file.|

## Remarks

If you want to fill the shape with small tiles of an image, use the **[UserTextured](Word.FillFormat.UserTextured.md)** method.


## Example

This example adds two rectangles to the active document. The rectangle on the left is filled with one large image of the picture in Tiles.bmp; the rectangle on the right is filled with many small tiles of the picture in Tiles.bmp.

```vb
Sub Pic() 
 With ActiveDocument.Shapes 
 .AddShape(msoShapeRectangle, 0, 0, 200, 100).Fill _ 
 .UserPicture "C:\Windows\Tiles.bmp" 
 .AddShape(msoShapeRectangle, 300, 0, 200, 100).Fill _ 
 .UserTextured "C:\Windows\Tiles.bmp" 
 End With 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]