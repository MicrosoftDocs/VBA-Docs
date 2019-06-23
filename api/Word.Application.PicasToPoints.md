---
title: Application.PicasToPoints method (Word)
keywords: vbawd10.chm158335349
f1_keywords:
- vbawd10.chm158335349
ms.prod: word
api_name:
- Word.Application.PicasToPoints
ms.assetid: ef812e9a-4bf5-b457-afa2-06371b411605
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PicasToPoints method (Word)

Converts a measurement from picas to points (1 pica = 12 points). Returns the converted measurement as a  **Single**.


## Syntax

_expression_. `PicasToPoints`( `_Picas_` )

_expression_ A variable that represents an **[Application](Word.Application.md)** object.  Optional.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Picas_|Required| **Single**|The pica value to be converted to points.|

## Return value

Single


## Example

This example adds line numbers to the active document and sets the distance between the line numbers and the document text to 4 picas.


```vb
With ActiveDocument.PageSetup.LineNumbering 
 .Active = True 
 .DistanceFromText = PicasToPoints(4) 
End With
```

This example sets the first-line indent for the selected paragraphs to 3 picas.




```vb
Selection.ParagraphFormat.FirstLineIndent = PicasToPoints(3)
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]