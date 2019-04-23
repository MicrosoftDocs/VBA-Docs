---
title: EmailOptions.AutoFormatAsYouTypeApplyClosings property (Word)
keywords: vbawd10.chm165347627
f1_keywords:
- vbawd10.chm165347627
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeApplyClosings
ms.assetid: b5be989e-09ff-455f-5d8a-638016512e3d
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.AutoFormatAsYouTypeApplyClosings property (Word)

 **True** for Microsoft Word to automatically apply the Closing style to letter closings as you type. Read/write **Boolean**.


## Syntax

_expression_. `AutoFormatAsYouTypeApplyClosings`

_expression_ Required. A variable that represents an '[EmailOptions](Word.EmailOptions.md)' collection.


## Example

This example sets Microsoft Word to automatically apply the Closing style to letter closings as you type.


```vb
Sub AutoClosings() 
 Options.AutoFormatAsYouTypeApplyClosings = True 
End Sub
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]