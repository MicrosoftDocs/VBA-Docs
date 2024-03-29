---
title: Graphic.Brightness property (Excel)
keywords: vbaxl10.chm694073
f1_keywords:
- vbaxl10.chm694073
api_name:
- Excel.Graphic.Brightness
ms.assetid: 42776335-6992-b37d-39a8-4a388b56da3e
ms.date: 04/26/2019
ms.localizationpriority: medium
---


# Graphic.Brightness property (Excel)

Returns or sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write **Single**.


## Syntax

_expression_.**Brightness**

_expression_ A variable that represents a **[Graphic](Excel.Graphic.md)** object.


## Example

This example sets the brightness for shape one on _myDocument_. Shape one must be either a picture or an OLE object.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).PictureFormat.Brightness = 0.3
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]