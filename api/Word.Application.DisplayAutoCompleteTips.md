---
title: Application.DisplayAutoCompleteTips property (Word)
keywords: vbawd10.chm158335068
f1_keywords:
- vbawd10.chm158335068
ms.prod: word
api_name:
- Word.Application.DisplayAutoCompleteTips
ms.assetid: 1ffcf473-d6f5-e2e7-c02c-0038b3fd3004
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DisplayAutoCompleteTips property (Word)

 **True** if Word displays tips that suggest text for completing words, dates, or phrases as you type. Read/write **Boolean**.


## Syntax

_expression_. `DisplayAutoCompleteTips`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example sets Word to display tips that suggest text for completing words, dates, or phrases as you type.


```vb
Application.DisplayAutoCompleteTips = True
```

This example returns the status of the Suggest the rest of the word or date with a tip as you type option on the AutoText tab in the AutoCorrect dialog box (Tools menu).




```vb
Dim blnTemp As Boolean 
 
blnTemp = Application.DisplayAutoCompleteTips
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]