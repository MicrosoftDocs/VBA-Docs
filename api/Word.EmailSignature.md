---
title: EmailSignature object (Word)
keywords: vbawd10.chm2524
f1_keywords:
- vbawd10.chm2524
ms.prod: word
api_name:
- Word.EmailSignature
ms.assetid: 9d641321-d52b-ab9a-4117-6f9e11dedbba
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailSignature object (Word)

Contains information about the email signatures used by Microsoft Word when you create and edit email messages and replies.


## Remarks

Use the  **EmailSignature** property to return the **EmailSignature** object.

This example changes the signatures Word appends to new outgoing email messages and email message replies.




```vb
With Application.EmailOptions.EmailSignature 
 .NewMessageSignature = "Signature1" 
 .ReplyMessageSignature = "Reply2" 
End With
```


> [!NOTE] 
> There is no EmailSignatures collection; each  **[EmailOptions](Word.EmailOptions.md)** object contains only one **EmailSignature** object.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]