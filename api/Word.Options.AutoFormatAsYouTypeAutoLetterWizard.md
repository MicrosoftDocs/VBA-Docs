---
title: Options.AutoFormatAsYouTypeAutoLetterWizard property (Word)
keywords: vbawd10.chm162988336
f1_keywords:
- vbawd10.chm162988336
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeAutoLetterWizard
ms.assetid: be49edd1-cb44-12d1-df43-ddaaddccef04
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatAsYouTypeAutoLetterWizard property (Word)

 **True** for Microsoft Word to automatically start the Letter Wizard when the user enters a letter salutation or closing. Read/write.


## Syntax

_expression_. `AutoFormatAsYouTypeAutoLetterWizard`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to automatically start the Letter Wizard when the user enters a letter salutation or closing.


```vb
Sub AutoLetterWizard() 
 Options.AutoFormatAsYouTypeAutoLetterWizard = True 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]