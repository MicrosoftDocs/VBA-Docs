---
title: Options.AutoFormatAsYouTypeDefineStyles property (Word)
keywords: vbawd10.chm162988302
f1_keywords:
- vbawd10.chm162988302
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeDefineStyles
ms.assetid: 16657544-0185-204f-1cee-b959c91956d5
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatAsYouTypeDefineStyles property (Word)

 **True** if Word automatically creates new styles based on manual formatting. Read/write **Boolean**.


## Syntax

_expression_. `AutoFormatAsYouTypeDefineStyles`

_expression_ A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets Word to automatically create styles as you type.


```vb
Options.AutoFormatAsYouTypeDefineStyles = True
```

This example returns the status of the Define styles based on your formatting option on the AutoFormat As You Type tab in the AutoCorrect dialog box (Tools menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeDefineStyles
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]