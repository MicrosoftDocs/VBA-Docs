---
title: PictureFormat.TransparencyColor property (Word)
keywords: vbawd10.chm164298859
f1_keywords:
- vbawd10.chm164298859
ms.prod: word
api_name:
- Word.PictureFormat.TransparencyColor
ms.assetid: 5b332d25-0aef-15c3-3826-322ea697522c
ms.date: 06/08/2017
localization_priority: Normal
---


# PictureFormat.TransparencyColor property (Word)

Returns or sets the transparent color for the specified picture as a red-green-blue (RGB) value. Read/write  **Long**.


## Syntax

_expression_.**TransparencyColor**

 _expression_ An expression that returns a '[PictureFormat](Word.PictureFormat.md)' object.


## Remarks

For this property to take effect, the **[TransparentBackground](Word.PictureFormat.TransparentBackground.md)** property must be set to **True**. Applies to bitmaps only.

If you want to be able to see through the transparent parts of the picture all the way to the objects behind the picture, you must set the **[Visible](Word.FillFormat.Visible.md)** property of the picture's **[FillFormat](Word.FillFormat.md)** object to **False**. If your picture has a transparent color and the **Visible** property of the picture's **FillFormat** object is set to **True**, the picture's fill will be visible through the transparent color, but objects behind the picture will be obscured.


## Example

This example sets the color returned by the RGB function as the transparent color for shape one in the active document. For the example to work, shape one must be a bitmap.


```vb
blueScreen = RGB(0, 0, 255) 
With ActiveDocument.Shapes(1) 
 With .PictureFormat 
 .TransparentBackground = True 
 .TransparencyColor = blueScreen 
 End With 
 .Fill.Visible = False 
End With
```


## See also


[PictureFormat Object](Word.PictureFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]