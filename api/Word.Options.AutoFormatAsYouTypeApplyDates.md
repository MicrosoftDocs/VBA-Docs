---
title: Options.AutoFormatAsYouTypeApplyDates property (Word)
keywords: vbawd10.chm162988330
f1_keywords:
- vbawd10.chm162988330
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeApplyDates
ms.assetid: b31f13fa-9a76-3a86-c4c2-4720fec1b66b
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatAsYouTypeApplyDates property (Word)

 **True** for Microsoft Word to automatically apply the Date style to dates as you type. Read/write.


## Syntax

_expression_. `AutoFormatAsYouTypeApplyDates`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to automatically apply the Date style to dates as you type.


```vb
Sub AutoApplyDates() 
 Options.AutoFormatAsYouTypeApplyDates = True 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]