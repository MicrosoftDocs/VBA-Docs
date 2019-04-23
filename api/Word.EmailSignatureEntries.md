---
title: EmailSignatureEntries object (Word)
keywords: vbawd10.chm2533
f1_keywords:
- vbawd10.chm2533
ms.prod: word
api_name:
- Word.EmailSignatureEntries
ms.assetid: 42a63f45-f989-be32-e75a-059c9a77c6f1
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailSignatureEntries object (Word)

A collection of  **[EmailSignatureEntry](Word.EmailSignatureEntry.md)** objects that represents all the email signature entries available to Word.


## Remarks

Use the  **[EmailSignatureEntries](Word.EmailSignature.EmailSignatureEntries.md)** property to return the **EmailSignatureEntries** collection. Use the **[Add](Word.EmailSignatureEntries.Add.md)** method of the **EmailSignatureEntries** object to add an email signature to Word. The following example creates a new email signature entry based on the author's name and a selection in the active document, and then it sets the new signature entry as the default email signature to use for new messages.


```vb
Sub NewEmailSignature() 
 With Application.EmailOptions.EmailSignature 
 .EmailSignatureEntries.Add "Jeff Smith", Selection.Range 
 .NewMessageSignature = "Jeff Smith" 
 End With 
End Sub
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]