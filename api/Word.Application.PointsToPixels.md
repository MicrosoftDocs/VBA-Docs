---
title: Application.PointsToPixels method (Word)
keywords: vbawd10.chm158335363
f1_keywords:
- vbawd10.chm158335363
ms.prod: word
api_name:
- Word.Application.PointsToPixels
ms.assetid: fc8eabb3-75f0-e456-bbd0-c17daa5ad1f3
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PointsToPixels method (Word)

Converts a measurement from points to pixels. Returns the converted measurement as a  **Single**.


## Syntax

_expression_. `PointsToPixels`( `_Points_` , `_fVertical_` )

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


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


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]