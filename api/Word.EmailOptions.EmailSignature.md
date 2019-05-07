---
title: EmailOptions.EmailSignature property (Word)
keywords: vbawd10.chm165347436
f1_keywords:
- vbawd10.chm165347436
ms.prod: word
api_name:
- Word.EmailOptions.EmailSignature
ms.assetid: 853e0b8d-8e25-4626-154f-1d634e485929
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.EmailSignature property (Word)

Returns an  **[EmailSignature](Word.EmailSignature.md)** object that represents the signatures Microsoft Word appends to outgoing email messages. Read-only.


## Syntax

_expression_. `EmailSignature`

_expression_ A variable that represents a '[EmailOptions](Word.EmailOptions.md)' object.


## Example

This example displays the signature Word appends to new outgoing email messages.


```vb
With Application.EmailOptions.EmailSignature 
 If .NewMessageSignature = "" Then 
 MsgBox "There is no signature for new " _ 
 & "email messages!" 
 Else 
 MsgBox "The signature for new email" _ 
 & "messages is: " & vbLf & vbLf _ 
 & .NewMessageSignature 
 End If 
End With
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]