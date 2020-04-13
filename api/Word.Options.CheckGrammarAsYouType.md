---
title: Options.CheckGrammarAsYouType property (Word)
keywords: vbawd10.chm162988309
f1_keywords:
- vbawd10.chm162988309
ms.prod: word
api_name:
- Word.Options.CheckGrammarAsYouType
ms.assetid: 11e4c676-bd8d-26e0-a0d4-74537508fc88
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.CheckGrammarAsYouType property (Word)

 **True** if Word checks grammar and marks errors automatically as you type. Read/write **Boolean**.


## Syntax

_expression_. `CheckGrammarAsYouType`

_expression_ A variable that represents a '[Options](Word.Options.md)' object.


## Remarks

This property marks grammatical errors, but to see them on screen, you must set the **[ShowGrammaticalErrors](Word.Document.ShowGrammaticalErrors.md)** property to **True**.


## Example

This example sets Word to check for grammatical errors as you type and to display any errors found in the active document.


```vb
Options.CheckGrammarAsYouType = True 
ActiveDocument.ShowGrammaticalErrors = True
```

This example returns the status of the **Check grammar as you type** option on the **Spelling & Grammar** tab in the **Options** dialog box (**Tools** menu).




```vb
Dim blnCheck As Boolean 
 
blnCheck = Options.CheckGrammarAsYouType
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]