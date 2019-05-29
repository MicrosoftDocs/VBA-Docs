---
title: Options.TabIndentKey property (Word)
keywords: vbawd10.chm162988104
f1_keywords:
- vbawd10.chm162988104
ms.prod: word
api_name:
- Word.Options.TabIndentKey
ms.assetid: 1edd2ffe-29ce-a4cc-6986-2f14ac03fb7a
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.TabIndentKey property (Word)

 **True** if the TAB and BACKSPACE keys can be used to increase and decrease, respectively, the left indent of paragraphs and if the BACKSPACE key can be used to change right-aligned paragraphs to centered paragraphs and centered paragraphs to left-aligned paragraphs. Read/write **Boolean**.


## Syntax

_expression_. `TabIndentKey`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets Word so that the TAB and BACKSPACE keys set the left indent instead of inserting and deleting tabs.


```vb
Options.TabIndentKey = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]