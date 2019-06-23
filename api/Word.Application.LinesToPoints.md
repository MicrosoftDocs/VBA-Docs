---
title: Application.LinesToPoints method (Word)
keywords: vbawd10.chm158335350
f1_keywords:
- vbawd10.chm158335350
ms.prod: word
api_name:
- Word.Application.LinesToPoints
ms.assetid: f146db0f-35f6-d25d-2674-e35a7c08801b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.LinesToPoints method (Word)

Converts a measurement from lines to points (1 line = 12 points). Returns the converted measurement as a  **Single**.


## Syntax

_expression_. `LinesToPoints`( `_Lines_` )

_expression_ A variable that represents an **[Application](Word.Application.md)** object.  Optional.


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


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]