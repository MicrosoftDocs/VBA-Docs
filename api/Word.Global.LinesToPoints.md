---
title: Global.LinesToPoints method (Word)
keywords: vbawd10.chm163119478
f1_keywords:
- vbawd10.chm163119478
ms.prod: word
api_name:
- Word.Global.LinesToPoints
ms.assetid: 3acbbbef-0aec-d6aa-138f-cdd1e79e7dc6
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.LinesToPoints method (Word)

Converts a measurement from lines to points (1 line = 12 points). Returns the converted measurement as a  **Single**.


## Syntax

_expression_. `LinesToPoints`( `_Lines_` )

_expression_ A variable that represents a '[Global](Word.Global.md)' object. Optional.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Lines_|Required| **Single**|The line value to be converted to points.|

## Return value

Single


## Example

This example sets the paragraph line spacing in the selection to three lines.


```vb
With Selection.ParagraphFormat 
 .LineSpacingRule = wdLineSpaceMultiple 
 .LineSpacing = LinesToPoints(3) 
End With
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]