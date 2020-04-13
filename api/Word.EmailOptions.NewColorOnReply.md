---
title: EmailOptions.NewColorOnReply property (Word)
keywords: vbawd10.chm165347444
f1_keywords:
- vbawd10.chm165347444
ms.prod: word
api_name:
- Word.EmailOptions.NewColorOnReply
ms.assetid: f7878b23-46a3-7950-7b45-28810de58f91
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.NewColorOnReply property (Word)

 **True** specifies whether a user needs to choose a new color for reply text when replying to email. Read/write **Boolean**.


## Syntax

_expression_. `NewColorOnReply`

 _expression_ An expression that returns an '[EmailOptions](Word.EmailOptions.md)' object.


## Remarks

Use the **NewColorOnReply** property if you want the reply text of email messages sent from Microsoft Word to be a different color than the original message.


## Example

This example checks to see if a user needs to choose a new color for email reply text, and if not, sets the reply font color to blue.


```vb
Sub NewColor() 
 With Application.EmailOptions 
 If .NewColorOnReply = False Then 
 .ReplyStyle.Font.Color = wdColorBlue 
 End If 
 End With 
End Sub
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]