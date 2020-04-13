---
title: EmailSignature.NewMessageSignature property (Word)
keywords: vbawd10.chm165412967
f1_keywords:
- vbawd10.chm165412967
ms.prod: word
api_name:
- Word.EmailSignature.NewMessageSignature
ms.assetid: fed9f151-47b8-3e76-1764-b6e80bdbfb5e
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailSignature.NewMessageSignature property (Word)

Returns or sets the signature that Microsoft Word appends to new email messages. Read/write  **String**.


## Syntax

_expression_. `NewMessageSignature`

 _expression_ An expression that returns an '[EmailSignature](Word.EmailSignature.md)' object.


## Remarks

When setting this property, you must use the name of an email signature that you have created in the **Email Options** dialog box, available from the **General** tab of the **Options** dialog box (**Tools** menu).


## Example

This example changes the signature Word appends to new outgoing email messages.


```vb
With Application.EmailOptions.EmailSignature 
 .NewMessageSignature = "Signature1" 
End With
```


## See also


[EmailSignature Object](Word.EmailSignature.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]