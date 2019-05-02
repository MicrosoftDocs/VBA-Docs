---
title: PictureFormat.Contrast property (Excel)
keywords: vbaxl10.chm113004
f1_keywords:
- vbaxl10.chm113004
ms.prod: excel
api_name:
- Excel.PictureFormat.Contrast
ms.assetid: 994cfca5-8ddb-d943-63c8-21abe8508de6
ms.date: 05/03/2019
localization_priority: Normal
---


# PictureFormat.Contrast property (Excel)

Returns or sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write **Single**.


## Syntax

_expression_.**Contrast**

 _expression_ An expression that returns a **[PictureFormat](Excel.PictureFormat.md)** object.


## Example

This example sets the contrast for shape one on _myDocument_. Shape one must be either a picture or an OLE object.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).PictureFormat.Contrast = 0.8
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]