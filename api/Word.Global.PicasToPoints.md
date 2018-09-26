---
title: Global.PicasToPoints Method (Word)
keywords: vbawd10.chm163119477
f1_keywords:
- vbawd10.chm163119477
ms.prod: word
api_name:
- Word.Global.PicasToPoints
ms.assetid: c1fb493b-d63d-484f-9d9b-c6781a0ff027
ms.date: 06/08/2017
---


# Global.PicasToPoints Method (Word)

Converts a measurement from picas to points (1 pica = 12 points). Returns the converted measurement as a  **Single** .


## Syntax

 _expression_. `PicasToPoints`( `_Picas_` )

 _expression_ A variable that represents a '[Global](Word.Global.md)' object. Optional.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Picas_|Required| **Single**|The pica value to be converted to points.|

### Return value

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


[Global Object](Word.Global.md)

