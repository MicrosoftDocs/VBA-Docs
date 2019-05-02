---
title: PictureFormat.IncrementBrightness method (Excel)
keywords: vbaxl10.chm113020
f1_keywords:
- vbaxl10.chm113020
ms.prod: excel
api_name:
- Excel.PictureFormat.IncrementBrightness
ms.assetid: 3f75ff17-6cd6-e397-468c-6bf0d1307578
ms.date: 05/03/2019
localization_priority: Normal
---


# PictureFormat.IncrementBrightness method (Excel)

Changes the brightness of the picture by the specified amount. Use the **[Brightness](Excel.PictureFormat.Brightness.md)** property to set the absolute brightness of the picture.


## Syntax

_expression_.**IncrementBrightness** (_Increment_)

_expression_ A variable that represents a **[PictureFormat](Excel.PictureFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how much to change the value of the **Brightness** property for the picture. A positive value makes the picture brighter; a negative value makes the picture darker.|

## Remarks

You cannot adjust the brightness of a picture past the upper or lower limit for the **Brightness** property. For example, if the **Brightness** property is initially set to 0.9 and you specify 0.3 for the _Increment_ argument, the resulting brightness level will be 1.0, which is the upper limit for the **Brightness** property, instead of 1.2.


## Example

This example creates a duplicate of shape one on _myDocument_, and then moves and darkens the duplicate. For the example to work, shape one must be either a picture or an OLE object.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Duplicate 
 .PictureFormat.IncrementBrightness -0.2 
 .IncrementLeft 50 
 .IncrementTop 50 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]