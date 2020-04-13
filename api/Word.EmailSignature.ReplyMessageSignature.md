---
title: EmailSignature.ReplyMessageSignature property (Word)
keywords: vbawd10.chm165412968
f1_keywords:
- vbawd10.chm165412968
ms.prod: word
api_name:
- Word.EmailSignature.ReplyMessageSignature
ms.assetid: 94e6bc68-8bf2-0c08-b361-1792eafb089d
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailSignature.ReplyMessageSignature property (Word)

Returns or sets the signature that Microsoft Word appends to email message replies. Read/write  **String**.


## Syntax

_expression_. `ReplyMessageSignature`

 _expression_ An expression that returns an '[EmailSignature](Word.EmailSignature.md)' object.


## Remarks

When setting this property, you must use the name of an email signature that you have created in the **Email Options** dialog box, available from the **General** tab of the **Options** dialog box (**Tools** menu).


## Example

This example changes the signature Word appends to email message replies.


```vb
With Application.EmailOptions.EmailSignature 
 .ReplyMessageSignature = "Reply2" 
End With
```


## See also


[EmailSignature Object](Word.EmailSignature.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]