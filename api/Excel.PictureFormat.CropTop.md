---
title: PictureFormat.CropTop property (Excel)
keywords: vbaxl10.chm113008
f1_keywords:
- vbaxl10.chm113008
ms.prod: excel
api_name:
- Excel.PictureFormat.CropTop
ms.assetid: adde9cc2-ca09-8494-d250-92a36dfa51e0
ms.date: 05/03/2019
localization_priority: Normal
---


# PictureFormat.CropTop property (Excel)

Returns or sets the number of [points](../language/glossary/vbe-glossary.md#point) that are cropped off the top of the specified picture or OLE object. Read/write **Single**.


## Syntax

_expression_.**CropTop**

_expression_ An expression that returns a **[PictureFormat](Excel.PictureFormat.md)** object.


## Remarks

Cropping is calculated relative to the original size of the picture. For example, if you insert a picture that is originally 100 points high, rescale it so that it's 200 points high, and then set the **CropTop** property to 50, 100 points (not 50) will be cropped off the top of your picture.


## Example

This example crops 20 points off the top of shape three on _myDocument_. For the example to work, shape three must be either a picture or an OLE object.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(3).PictureFormat.CropTop = 20
```

<br/>

This example allows you to specify the percentage that you want to crop off the top of the selected shape, regardless of whether the shape has been scaled. For the example to work, the selected shape must be either a picture or an OLE object.

```vb
percentToCrop = InputBox( _ 
 "What percentage do you want to crop" & _ 
 " off the top of this picture?") 
Set shapeToCrop = ActiveWindow.Selection.ShapeRange(1) 
With shapeToCrop.Duplicate 
 .ScaleHeight 1, True 
 origHeight = .Height 
 .Delete 
End With 
cropPoints = origHeight * percentToCrop / 100 
shapeToCrop.PictureFormat.CropTop = cropPoints
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]