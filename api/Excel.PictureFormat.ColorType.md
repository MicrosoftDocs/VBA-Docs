---
title: PictureFormat.ColorType property (Excel)
keywords: vbaxl10.chm113003
f1_keywords:
- vbaxl10.chm113003
ms.prod: excel
api_name:
- Excel.PictureFormat.ColorType
ms.assetid: 6c183163-8fbd-3a0f-b087-05d8d2cdbfd5
ms.date: 05/03/2019
localization_priority: Normal
---


# PictureFormat.ColorType property (Excel)

Returns or sets the type of color transformation applied to the specified picture or OLE object. Read/write.


## Syntax

_expression_.**ColorType**

_expression_ An expression that returns a **[PictureFormat](Excel.PictureFormat.md)** object.


## Example

This example sets the color transformation to grayscale through **[MsoPictureColorType](Office.MsoPictureColorType.md)** for shape one on _myDocument_. Shape one must be either a picture or an OLE object.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).PictureFormat.ColorType = msoPictureGrayScale
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]