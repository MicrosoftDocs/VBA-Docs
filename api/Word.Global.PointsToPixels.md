---
title: Global.PointsToPixels method (Word)
keywords: vbawd10.chm163119489
f1_keywords:
- vbawd10.chm163119489
ms.prod: word
api_name:
- Word.Global.PointsToPixels
ms.assetid: e119ddf1-851c-2870-73f4-52da1d17c035
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.PointsToPixels method (Word)

Converts a measurement from points to pixels. Returns the converted measurement as a  **Single**.


## Syntax

_expression_. `PointsToPixels`( `_Points_` , `_fVertical_` )

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Points_|Required| **Single**|The point value to be converted to pixels.|
| _fVertical_|Optional| **Variant**| **True** to return the result as vertical pixels; **False** to return the result as horizontal pixels.|

## Return value

Single


## Example

This example displays the height and width in pixels of an object measured in points.


```vb
MsgBox "180x120 points is equivalent to " _ 
 & PointsToPixels(180, False) & "x" _ 
 & PointsToPixels(120, True) _ 
 & " pixels on this display."
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]