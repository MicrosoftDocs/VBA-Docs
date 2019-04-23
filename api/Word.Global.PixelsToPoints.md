---
title: Global.PixelsToPoints method (Word)
keywords: vbawd10.chm163119490
f1_keywords:
- vbawd10.chm163119490
ms.prod: word
api_name:
- Word.Global.PixelsToPoints
ms.assetid: 671b06c5-c54f-417f-557b-53ea9fee1480
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.PixelsToPoints method (Word)

Converts a measurement from pixels to points. Returns the converted measurement as a  **Single**.


## Syntax

_expression_. `PixelsToPoints`( `_Pixels_` , `_fVertical_` )

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pixels_|Required| **Single**|The pixel value to be converted to points.|
| _fVertical_|Optional| **Variant**| **True** to convert vertical pixels; **False** to convert horizontal pixels.|

## Return value

Single


## Example

This example displays the height and width in points of an object measured in pixels.


```vb
MsgBox "320x240 pixels is equivalent to " _ 
 & PixelsToPoints(320, False) & "x" _ 
 & PixelsToPoints(240, True) _ 
 & " points on this display."
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]