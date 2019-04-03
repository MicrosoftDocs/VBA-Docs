---
title: EmailSignatureEntry object (Word)
keywords: vbawd10.chm2534
f1_keywords:
- vbawd10.chm2534
ms.prod: word
api_name:
- Word.EmailSignatureEntry
ms.assetid: a8cf11de-7a46-c609-3cd7-508e9ef91e09
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailSignatureEntry object (Word)

Represents a single email signature entry. The  **EmailSignatureEntry** object is a member of the **[EmailSignatureEntries](Word.EmailSignatureEntries.md)** collection. The **EmailSignatureEntries** collection contains all the email signature entries available to Word.


## Remarks

Use  **EmailSignatureEntries** (Index), where Index is the email signature entry name or item number, to return a single **EmailSignatureEntry** object. You must match exactly the spelling (but not necessarily the capitalization) of the name. The following example uses the **[Delete](Word.EmailSignatureEntry.Delete.md)** method to delete the signature entry named "Jeff Smith."


```vb
Sub DeleteSignature() 
 Application.EmailOptions.EmailSignature _ 
 .EmailSignatureEntries("jeff smith").Delete 
End Sub
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]