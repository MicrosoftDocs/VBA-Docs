---
title: EmailOptions object (Word)
ms.prod: word
api_name:
- Word.EmailOptions
ms.assetid: 41fefa03-c993-e218-0f92-0cf30c0bfbd4
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions object (Word)

Contains global application-level attributes used by Microsoft Word when you create and edit email messages and replies.


## Remarks

Use the **[EmailOptions](Word.Application.EmailOptions.md)** property to return the **EmailOptions** object.

This example changes the font color of the default style used to compose new email messages.




```vb
Application.EmailOptions.ComposeStyle.Font.Color = _ 
 wdColorBrightGreen
```

This example sets Word to mark comments in email messages with the initials "WK."




```vb
Application.EmailOptions.MarkCommentsWith = "WK" 
Application.EmailOptions.MarkComments = True
```

This example changes the signatures Word appends to new outgoing email messages and email message replies.




```vb
With Application.EmailOptions.EmailSignature 
 .NewMessageSignature = "Signature1" 
 .ReplyMessageSignature = "Reply2" 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]