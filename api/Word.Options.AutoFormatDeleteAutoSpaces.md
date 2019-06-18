---
title: Options.AutoFormatDeleteAutoSpaces property (Word)
keywords: vbawd10.chm162988328
f1_keywords:
- vbawd10.chm162988328
ms.prod: word
api_name:
- Word.Options.AutoFormatDeleteAutoSpaces
ms.assetid: 45f56b46-bdb5-972b-d4f7-ba736a80d4c1
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatDeleteAutoSpaces property (Word)

 **True** if spaces inserted between Japanese and Latin text will be deleted when Microsoft Word formats a document or range automatically. Read/write **Boolean**.


## Syntax

_expression_. `AutoFormatDeleteAutoSpaces`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to automatically delete spaces between Japanese and Latin text, and then it formats the current selection.


```vb
Options.AutoFormatDeleteAutoSpaces = True 
Selection.Range.AutoFormat
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]