---
title: Graphic.CropBottom property (Excel)
keywords: vbaxl10.chm694076
f1_keywords:
- vbaxl10.chm694076
ms.prod: excel
api_name:
- Excel.Graphic.CropBottom
ms.assetid: ae0ab176-81c4-73ef-bfdb-e08c078dfc9e
ms.date: 04/26/2019
localization_priority: Normal
---


# Graphic.CropBottom property (Excel)

Returns or sets the number of [points](../language/glossary/vbe-glossary.md#point) that are cropped off the bottom of the specified picture or OLE object. Read/write **Single**.


## Syntax

_expression_.**CropBottom**

_expression_ An expression that returns a **[Graphic](Excel.Graphic.md)** object.


## Remarks

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points high, rescale it so that it's 200 points high, and then set the **CropBottom** property to 50, 100 points (not 50) will be cropped off the bottom of your picture.


## Example

This example crops 20 points off the bottom of shape three on _myDocument_. For the example to work, shape three must be either a picture or an OLE object.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(3).PictureFormat.CropBottom = 20
```

<br/>

Using the following example, you can specify the percentage that you want to crop off the bottom of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.

```vb
percentToCrop = InputBox( _ 
 "What percentage do you want to crop off" & _ 
 " the bottom of this picture?") 
Set shapeToCrop = ActiveWindow.Selection.ShapeRange(1) 
With shapeToCrop.Duplicate 
 .ScaleHeight 1, True 
 origHeight = .Height 
 .Delete 
End With 
cropPoints = origHeight * percentToCrop / 100 
shapeToCrop.PictureFormat.CropBottom = cropPoints
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]