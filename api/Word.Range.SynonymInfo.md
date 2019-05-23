---
title: Range.SynonymInfo property (Word)
keywords: vbawd10.chm157155483
f1_keywords:
- vbawd10.chm157155483
ms.prod: word
api_name:
- Word.Range.SynonymInfo
ms.assetid: b63d2a0b-baa1-306d-10ee-72223099a9f2
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.SynonymInfo property (Word)

Returns a  **SynonymInfo** object that contains information from the thesaurus on synonyms, antonyms, or related words and expressions for the contents of a range.


## Syntax

_expression_. `SynonymInfo`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Example

This example returns a list of synonyms for the selection's first meaning.


```vb
Slist = Selection.Range.SynonymInfo.SynonymList(Meaning:=1) 
For i = 1 To UBound(Slist) 
 Msgbox Slist(i) 
Next i
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]