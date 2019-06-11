---
title: PictureFormat.IncrementBrightness method (Publisher)
keywords: vbapb10.chm3604496
f1_keywords:
- vbapb10.chm3604496
ms.prod: publisher
api_name:
- Publisher.PictureFormat.IncrementBrightness
ms.assetid: 912fd08e-bbb3-bf98-b0da-7128926f3409
ms.date: 06/12/2019
localization_priority: Normal
---


# PictureFormat.IncrementBrightness method (Publisher)

Changes the brightness of the picture by the specified amount.


## Syntax

_expression_.**IncrementBrightness** (_Increment_)

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Increment_|Required| **Single**|Specifies how much to change the value of the **[Brightness](Publisher.PictureFormat.Brightness.md)** property for the picture. A positive value makes the picture brighter; a negative value makes the picture darker. Valid values are between - 1 and 1.|

## Remarks

You cannot adjust the brightness of a picture past the upper or lower limit for the **Brightness** property. For example, if the **Brightness** property is initially set to 0.9 and you specify 0.3 for the _Increment_ argument, the resulting brightness level will be 1.0, which is the upper limit for the **Brightness** property, instead of 1.2.

Use the **Brightness** property to set the absolute brightness of the picture.


## Example

This example creates a duplicate of the first shape in the active publication and then moves and darkens the duplicate. For the example to work, the shape must be either a picture or an OLE object representing a picture.

```vb
With ActiveDocument.Pages(1).Shapes(1).Duplicate 
 .PictureFormat.IncrementBrightness Increment:=-0.2 
 .IncrementLeft Increment:=50 
 .IncrementTop Increment:=50 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]