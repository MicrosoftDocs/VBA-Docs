---
title: Options.AutoFormatAsYouTypeFormatListItemBeginning property (Word)
keywords: vbawd10.chm162988301
f1_keywords:
- vbawd10.chm162988301
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeFormatListItemBeginning
ms.assetid: 7fc572d7-59f2-cb23-4609-c5ba6af9065c
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatAsYouTypeFormatListItemBeginning property (Word)

 **True** if Word repeats character formatting applied to the beginning of a list item to the next list item. Read/write **Boolean**.


## Syntax

_expression_. `AutoFormatAsYouTypeFormatListItemBeginning`

_expression_ A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets Word to automatically repeat character formatting at the beginning of list items.


```vb
Options.AutoFormatAsYouTypeFormatListItemBeginning = True
```

This example returns the status of the Format beginning of list item like the one before it option in the AutoFormat As You Type tab in the AutoCorrect dialog box (Options menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = _ 
 Options.AutoFormatAsYouTypeFormatListItemBeginning
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]