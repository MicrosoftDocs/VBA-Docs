---
title: ListFormat.RemoveNumbers method (Word)
keywords: vbawd10.chm163578041
f1_keywords:
- vbawd10.chm163578041
ms.prod: word
api_name:
- Word.ListFormat.RemoveNumbers
ms.assetid: 80c0e408-683d-4639-733d-843d5fd323e2
ms.date: 06/08/2017
localization_priority: Normal
---


# ListFormat.RemoveNumbers method (Word)

Removes numbers or bullets from the specified list.


## Syntax

_expression_. `RemoveNumbers`( `_NumberType_` )

_expression_ A variable that represents a '[ListFormat](Word.ListFormat.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NumberType_|Optional| **[WdNumberType](Word.WdNumberType.md)**| The type of number to be removed.|

## Example

This example removes the bullets or numbers from any numbered paragraphs in the selection.


```vb
Selection.Range.ListFormat.RemoveNumbers
```

This example removes the LISTNUM fields from the selection.




```vb
Selection.Range.ListFormat.RemoveNumbers wdNumberListNum
```


## See also


[ListFormat Object](Word.ListFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]