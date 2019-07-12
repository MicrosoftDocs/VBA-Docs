---
title: PictureFormat.IncrementBrightness method (PowerPoint)
keywords: vbapp10.chm551002
f1_keywords:
- vbapp10.chm551002
ms.prod: powerpoint
api_name:
- PowerPoint.PictureFormat.IncrementBrightness
ms.assetid: 4237d547-2c8b-9ed2-f131-6a4fb52ee0a2
ms.date: 06/08/2017
localization_priority: Normal
---


# PictureFormat.IncrementBrightness method (PowerPoint)

Changes the brightness of the picture by the specified amount. 


## Syntax

_expression_.**IncrementBrightness** (_Increment_)

_expression_ A variable that represents an [PictureFormat](PowerPoint.PictureFormat.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how much to change the value of the  **Brightness** property for the picture. A positive value makes the picture brighter; a negative value makes the picture darker.|

## Remarks

Use the  **[Brightness](PowerPoint.PictureFormat.Brightness.md)** property to set the absolute brightness of the picture.

You cannot adjust the brightness of a picture past the upper or lower limit for the  **Brightness** property. For example, if the **Brightness** property is initially set to 0.9 and you specify 0.3 for the Increment argument, the resulting brightness level will be 1.0, which is the upper limit for the **Brightness** property, instead of 1.2.


## Example

This example creates a duplicate of shape one on _myDocument_ and then moves and darkens the duplicate. For the example to work, shape one must be either a picture or an OLE object.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).Duplicate

    .PictureFormat.IncrementBrightness -0.2

    .IncrementLeft 50

    .IncrementTop 50

End With
```


## See also


[PictureFormat Object](PowerPoint.PictureFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]