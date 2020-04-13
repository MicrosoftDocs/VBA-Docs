---
title: Range.CheckSynonyms method (Word)
keywords: vbawd10.chm157155508
f1_keywords:
- vbawd10.chm157155508
ms.prod: word
api_name:
- Word.Range.CheckSynonyms
ms.assetid: e28026bf-aa5e-8cf4-e765-7350afd57741
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.CheckSynonyms method (Word)

Displays the **Thesaurus** dialog box, which lists alternative word choices, or synonyms, for the text in the specified range.


## Syntax

_expression_. `CheckSynonyms`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Example

This example displays the **Thesaurus** dialog box with synonyms for the selected text.


```vb
Selection.Range.CheckSynonyms
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]