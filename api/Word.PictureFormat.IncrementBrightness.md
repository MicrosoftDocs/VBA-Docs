---
title: PictureFormat.IncrementBrightness method (Word)
keywords: vbawd10.chm164298762
f1_keywords:
- vbawd10.chm164298762
ms.prod: word
api_name:
- Word.PictureFormat.IncrementBrightness
ms.assetid: 2bce8316-c15c-e5b9-9f04-1095ccaa7126
ms.date: 06/08/2017
localization_priority: Normal
---


# PictureFormat.IncrementBrightness method (Word)

Changes the brightness of the picture by the specified amount.


## Syntax

_expression_.**IncrementBrightness** (_Increment_)

_expression_ Required. A variable that represents a '[PictureFormat](Word.PictureFormat.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how much to change the value of the **Brightness** property for the picture. A positive value makes the picture brighter; a negative value makes the picture darker.|

## Remarks

You cannot adjust the brightness of a picture past the upper or lower limit for the **Brightness** property. For example, if the **Brightness** property is initially set to 0.9 and you specify 0.3 for the Increment argument, the resulting brightness level will be 1.0, which is the upper limit for the **Brightness** property, instead of 1.2.

Use the **[Brightness](Word.PictureFormat.Brightness.md)** property to set the absolute brightness of the picture.


## Example

This example creates a duplicate of the first shape on the active document and then moves and darkens the duplicate. For the example to work, the first shape must be either a picture or an OLE object.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes(1).Duplicate 
 .PictureFormat.IncrementBrightness -0.2 
 .IncrementLeft 50 
 .IncrementTop 50 
End With
```


## See also


[PictureFormat Object](Word.PictureFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]