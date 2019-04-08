---
title: EmailOptions.TabIndentKey property (Word)
keywords: vbawd10.chm165347637
f1_keywords:
- vbawd10.chm165347637
ms.prod: word
api_name:
- Word.EmailOptions.TabIndentKey
ms.assetid: 48b79b45-5bc6-f253-acef-96f80cc68e1e
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.TabIndentKey property (Word)

 **True** if the TAB and BACKSPACE keys can be used to increase and decrease, respectively, the left indent of paragraphs and if the BACKSPACE key can be used to change right-aligned paragraphs to centered paragraphs and centered paragraphs to left-aligned paragraphs. Read/write **Boolean**.


## Syntax

_expression_. `TabIndentKey`

_expression_ Required. A variable that represents an '[EmailOptions](Word.EmailOptions.md)' collection.


## Example

This example sets Word so that the TAB and BACKSPACE keys set the left indent instead of inserting and deleting tabs.


```vb
Options.TabIndentKey = True
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]