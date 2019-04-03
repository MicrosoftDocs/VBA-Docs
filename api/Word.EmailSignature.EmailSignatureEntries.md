---
title: EmailSignature.EmailSignatureEntries property (Word)
keywords: vbawd10.chm165412969
f1_keywords:
- vbawd10.chm165412969
ms.prod: word
api_name:
- Word.EmailSignature.EmailSignatureEntries
ms.assetid: 8b5a2f6a-d9fe-5f92-d93d-a59e67ee7100
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailSignature.EmailSignatureEntries property (Word)

Returns an  **[EmailSignatureEntries](Word.EmailSignatureEntries.md)** object that represents the email signature entries in Microsoft Word. Read-only.


## Syntax

_expression_. `EmailSignatureEntries`

 _expression_ An expression that returns an '[EmailSignature](Word.EmailSignature.md)' object.


## Remarks

An email signature is standard text that ends an email message, such as your name and telephone number. Use the  **EmailSignatureEntries** property to create and manage a collection of email signatures that Word will use when creating email messages.


## Example

This example creates a new signature entry based on the author's name and the selection in the active document.


```vb
Sub NewSignature() 
 Application.EmailOptions.EmailSignature _ 
 .EmailSignatureEntries.Add _ 
 Name:=ActiveDocument.BuiltInDocumentProperties("Author"), _ 
 Range:=Selection.Range 
End Sub
```


## See also


[EmailSignature Object](Word.EmailSignature.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]