---
title: Selection.LanguageIDOther property (Word)
keywords: vbawd10.chm158662811
f1_keywords:
- vbawd10.chm158662811
ms.prod: word
api_name:
- Word.Selection.LanguageIDOther
ms.assetid: 197474ff-8d79-b48f-e1bf-ac2f164e70e3
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.LanguageIDOther property (Word)

Returns or sets the language for the specified object. Read/write  **WdLanguageID**.


## Syntax

_expression_. `LanguageIDOther`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

This is the recommended way to return or set the language of Latin text in a document created in a right-to-left language version of Microsoft Word.


## Example

This example sets the language of the selection to French.


```vb
Selection.LanguageIDOther = wdFrench
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]