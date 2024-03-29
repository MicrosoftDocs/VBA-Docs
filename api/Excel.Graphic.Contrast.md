---
title: Graphic.Contrast property (Excel)
keywords: vbaxl10.chm694075
f1_keywords:
- vbaxl10.chm694075
api_name:
- Excel.Graphic.Contrast
ms.assetid: 9715ee08-2d9b-1a5c-1fe9-3b5a73991668
ms.date: 04/26/2019
ms.localizationpriority: medium
---


# Graphic.Contrast property (Excel)

Returns or sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write **Single**.


## Syntax

_expression_.**Contrast**

_expression_ An expression that returns a **[Graphic](Excel.Graphic.md)** object.


## Example

This example sets the contrast for shape one on _myDocument_. Shape one must be either a picture or an OLE object.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).PictureFormat.Contrast = 0.8
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]