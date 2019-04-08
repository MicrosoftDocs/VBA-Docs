---
title: Document.GetLetterContent method (Word)
keywords: vbawd10.chm158007420
f1_keywords:
- vbawd10.chm158007420
ms.prod: word
api_name:
- Word.Document.GetLetterContent
ms.assetid: ab0d9fa4-b193-6a7f-641d-d6f971b37457
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.GetLetterContent method (Word)

Retrieves letter elements from the specified document and returns a  **[LetterContent](Word.LetterContent.md)** object.


## Syntax

_expression_. `GetLetterContent`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Return value

LetterContent


## Example

This example displays the salutation and recipient name from the letter in the active document.


```vb
MsgBox ActiveDocument.GetLetterContent.Salutation _ 
 & ActiveDocument.GetLetterContent.RecipientName
```

This example retrieves letter elements from the active document, changes the list of carbon copy (CC) recipients by setting the CClist property, and then uses the SetLetterContent method to update the active document to reflect the changes.




```vb
Set myLetterContent = ActiveDocument.GetLetterContent 
With myLetterContent 
 .CCList = "J. Burns, L. Scarpaczyk, K. Wong" 
 .RecipientName = "Amy Anderson" 
 .RecipientAddress = "123 Main" & vbCr & "Bellevue, WA 98004" 
 .LetterStyle = wdFullBlock 
End With 
ActiveDocument.SetLetterContent LetterContent:=myLetterContent
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]