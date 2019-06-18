---
title: Options.SuggestFromMainDictionaryOnly property (Word)
keywords: vbawd10.chm162988314
f1_keywords:
- vbawd10.chm162988314
ms.prod: word
api_name:
- Word.Options.SuggestFromMainDictionaryOnly
ms.assetid: d9ac9107-bf66-8f47-1101-6db4d6ec0364
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.SuggestFromMainDictionaryOnly property (Word)

 **True** if Microsoft Word draws spelling suggestions from the main dictionary only. Read/write **Boolean**.


## Syntax

_expression_. `SuggestFromMainDictionaryOnly`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

This property returns  **False** if it draws spelling suggestions from the main dictionary and any custom dictionaries that have been added.


## Example

This example sets Word to suggest words from the main dictionary only, and then it checks spelling in the active document.


```vb
Options.SuggestFromMainDictionaryOnly = True 
ActiveDocument.CheckSpelling
```

This example returns the current status of the Suggest from main dictionary only option on the Spelling & Grammar tab in the Options dialog box (Tools menu).




```vb
temp = Options.SuggestFromMainDictionaryOnly
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]