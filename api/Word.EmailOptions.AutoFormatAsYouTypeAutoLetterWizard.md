---
title: EmailOptions.AutoFormatAsYouTypeAutoLetterWizard property (Word)
keywords: vbawd10.chm165347632
f1_keywords:
- vbawd10.chm165347632
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeAutoLetterWizard
ms.assetid: 3a11e89f-7f02-e20c-4dcb-0bcf2724c043
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.AutoFormatAsYouTypeAutoLetterWizard property (Word)

 **True** for Microsoft Word to automatically start the Letter Wizard when the user enters a letter salutation or closing. Read/write.


## Syntax

 _expression_. `AutoFormatAsYouTypeAutoLetterWizard`

 _expression_ Required. A variable that represents an '[EmailOptions](Word.EmailOptions.md)' collection.


## Example

This example sets Microsoft Word to automatically start the Letter Wizard when the user enters a letter salutation or closing.


```vb
Sub AutoLeterWizard() 
 Options.AutoFormatAsYouTypeAutoLetterWizard = True 
End Sub
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

