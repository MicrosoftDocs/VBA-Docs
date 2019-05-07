---
title: Global.MillimetersToPoints method (Word)
keywords: vbawd10.chm163119476
f1_keywords:
- vbawd10.chm163119476
ms.prod: word
api_name:
- Word.Global.MillimetersToPoints
ms.assetid: c221d455-bcb1-f0eb-a658-53db12e06284
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.MillimetersToPoints method (Word)

Converts a measurement from millimeters to points (1 mm = 2.85 points). Returns the converted measurement as a  **Single**.


## Syntax

_expression_. `MillimetersToPoints`( `_Millimeters_` )

_expression_ A variable that represents a '[Global](Word.Global.md)' object. Optional.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Millimeters_|Required| **Single**|The millimeter value to be converted to points.|

## Return value

Single


## Example

This example sets the hyphenation zone in the active document to 8.8 millimeters.


```vb
ActiveDocument.HyphenationZone = MillimetersToPoints(8.8)
```

This example expands the spacing of the selected characters to 2.8 points.




```vb
Selection.Font.Spacing = MillimetersToPoints(1)
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]