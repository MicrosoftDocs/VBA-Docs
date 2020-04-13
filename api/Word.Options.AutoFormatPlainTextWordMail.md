---
title: Options.AutoFormatPlainTextWordMail property (Word)
keywords: vbawd10.chm162988303
f1_keywords:
- vbawd10.chm162988303
ms.prod: word
api_name:
- Word.Options.AutoFormatPlainTextWordMail
ms.assetid: 87b5f068-772c-e37d-9370-377849138d07
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatPlainTextWordMail property (Word)

 **True** if Word automatically formats plain-text email messages when you open them in Word. Read/write **Boolean**.


## Syntax

_expression_. `AutoFormatPlainTextWordMail`

_expression_ A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets Word to automatically format any plain-text email messages that are opened.


```vb
Options.AutoFormatPlainTextWordMail = True
```

This example returns the status of the **Plain text WordMail documents** option on the **AutoFormat** tab in the **AutoCorrect** dialog box (**Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatPlainTextWordMail
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]