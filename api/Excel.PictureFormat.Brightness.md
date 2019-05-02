---
title: PictureFormat.Brightness property (Excel)
keywords: vbaxl10.chm113002
f1_keywords:
- vbaxl10.chm113002
ms.prod: excel
api_name:
- Excel.PictureFormat.Brightness
ms.assetid: f17ee171-47da-c982-2f48-9ee333193add
ms.date: 05/03/2019
localization_priority: Normal
---


# PictureFormat.Brightness property (Excel)

Returns or sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write **Single**.


## Syntax

_expression_.**Brightness**

_expression_ A variable that represents a **[PictureFormat](Excel.PictureFormat.md)** object.


## Example

This example sets the brightness for shape one on _myDocument_. Shape one must be either a picture or an OLE object.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).PictureFormat.Brightness = 0.3
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]