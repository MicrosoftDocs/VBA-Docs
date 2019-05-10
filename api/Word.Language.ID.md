---
title: Language.ID property (Word)
keywords: vbawd10.chm158138378
f1_keywords:
- vbawd10.chm158138378
ms.prod: word
api_name:
- Word.Language.ID
ms.assetid: 8af15ba5-19f0-2a65-e44a-a9fed55f8239
ms.date: 06/08/2017
localization_priority: Normal
---


# Language.ID property (Word)

Returns a number that identifies the specified language. Read-only  **WdLanguageID**.


## Syntax

_expression_.**ID**

_expression_ Required. A variable that represents a '[Language](Word.Language.md)' object.


## Example

This example formats the selection with the Icelandic proofing tools language.


```vb
Selection.LanguageID = Languages("Icelandic").ID
```


## See also


[Language Object](Word.Language.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]