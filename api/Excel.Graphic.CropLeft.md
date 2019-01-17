---
title: Graphic.CropLeft property (Excel)
keywords: vbaxl10.chm694077
f1_keywords:
- vbaxl10.chm694077
ms.prod: excel
api_name:
- Excel.Graphic.CropLeft
ms.assetid: decebec1-af4a-2bb1-62b5-d90674b5b338
ms.date: 06/08/2017
localization_priority: Normal
---


# Graphic.CropLeft property (Excel)

Returns or sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/write  **Single**.


## Syntax

_expression_. `CropLeft`

 _expression_ An expression that returns a [Graphic](Excel.Graphic.md) object.


## Remarks

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points wide, rescale it so that it's 200 points wide, and then set the  **CropLeft** property to 50, 100 points (not 50) will be cropped off the left side of your picture.


## Example

This example crops 20 points off the left side of shape three on  `myDocument`. For the example to work, shape three must be either a picture or an OLE object.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(3).PictureFormat.CropLeft = 20
```

Using this example, you can specify the percentage you want to crop off the left side of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.




```vb
percentToCrop = InputBox( _ 
 "What percentage do you want to crop" & _ 
 " off the left of this picture?") 
Set shapeToCrop = ActiveWindow.Selection.ShapeRange(1) 
With shapeToCrop.Duplicate 
 .ScaleWidth 1, True 
 origWidth = .Width 
 .Delete 
End With 
cropPoints = origWidth * percentToCrop / 100 
shapeToCrop.PictureFormat.CropLeft = cropPoints
```


## See also


[Graphic Object](Excel.Graphic.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]