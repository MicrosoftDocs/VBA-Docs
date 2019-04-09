---
title: EmailOptions.AutoFormatAsYouTypeApplyDates property (Word)
keywords: vbawd10.chm165347626
f1_keywords:
- vbawd10.chm165347626
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeApplyDates
ms.assetid: e600d058-9864-84f7-7908-62ffe38d350a
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.AutoFormatAsYouTypeApplyDates property (Word)

 **True** for Microsoft Word to automatically apply the Date style to dates as you type. Read/write.


## Syntax

_expression_. `AutoFormatAsYouTypeApplyDates`

_expression_ Required. A variable that represents an '[EmailOptions](Word.EmailOptions.md)' collection.


## Example

This example sets Microsoft Word to automatically apply the Date style to dates as you type.


```vb
Sub AutoApplyDates() 
 Options.AutoFormatAsYouTypeApplyDates = True 
End Sub
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]