---
title: Options.AutoFormatApplyOtherParas property (Word)
keywords: vbawd10.chm162988285
f1_keywords:
- vbawd10.chm162988285
ms.prod: word
api_name:
- Word.Options.AutoFormatApplyOtherParas
ms.assetid: b6204429-d883-2235-f8c2-03e5d433c863
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatApplyOtherParas property (Word)

 **True** if styles are automatically applied to paragraphs that aren't headings or list items when Word formats a document or range automatically. Read/write **Boolean**.


## Syntax

_expression_. `AutoFormatApplyOtherParas`

_expression_ A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example automatically applies styles to paragraphs in the current selection.


```vb
Options.AutoFormatApplyOtherParas = True 
Selection.Range.AutoFormat
```

This example returns the status of the **Other paragraphs** option on the **AutoFormat** tab in the **AutoCorrect** dialog box (**Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatApplyOtherParas
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]