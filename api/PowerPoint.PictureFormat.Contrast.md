---
title: PictureFormat.Contrast property (PowerPoint)
keywords: vbapp10.chm551006
f1_keywords:
- vbapp10.chm551006
ms.prod: powerpoint
api_name:
- PowerPoint.PictureFormat.Contrast
ms.assetid: 19e2a7d2-59c3-e3d7-3770-0cbecdba2550
ms.date: 06/08/2017
localization_priority: Normal
---


# PictureFormat.Contrast property (PowerPoint)

Returns or sets the contrast for the specified picture or OLE object.


## Syntax

_expression_.**Contrast**

_expression_ A variable that represents a [PictureFormat](PowerPoint.PictureFormat.md) object.


## Return value

Single


## Remarks

The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write.


## Example

This example sets the contrast for shape one on _myDocument_. Shape one must be either a picture or an OLE object.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).PictureFormat.Contrast = 0.8
```


## See also


[PictureFormat Object](PowerPoint.PictureFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]