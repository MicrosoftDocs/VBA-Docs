---
title: EmailOptions.AutoFormatAsYouTypeApplyFirstIndents property (Word)
keywords: vbawd10.chm165347625
f1_keywords:
- vbawd10.chm165347625
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeApplyFirstIndents
ms.assetid: a05e77d8-9280-7754-e842-6fe3ae66eaa9
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.AutoFormatAsYouTypeApplyFirstIndents property (Word)

 **True** for Microsoft Word to automatically replace a space entered at the beginning of a paragraph with a first-line indent. Read/write.


## Syntax

 _expression_. `AutoFormatAsYouTypeApplyFirstIndents`

 _expression_ Required. A variable that represents an '[EmailOptions](Word.EmailOptions.md)' collection.


## Example

This example sets Microsoft Word to automatically replace a space entered at the beginning of a paragraph with a first-line indent as you type.


```vb
Sub ApplyFirstIndents() 
 Options.AutoFormatAsYouTypeApplyFirstIndents = True 
End Sub
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

