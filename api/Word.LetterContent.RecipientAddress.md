---
title: LetterContent.RecipientAddress property (Word)
keywords: vbawd10.chm161546349
f1_keywords:
- vbawd10.chm161546349
ms.prod: word
api_name:
- Word.LetterContent.RecipientAddress
ms.assetid: bcfbc400-0db7-0c86-5cb7-2a67a8ef9513
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.RecipientAddress property (Word)

Returns or sets the mailing address of the person who'll be receiving the letter created by the Letter Wizard. Read/write  **String**.


## Syntax

_expression_. `RecipientAddress`

 _expression_ An expression that returns a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example creates a new **LetterContent** object, sets several properties (including the recipient address), and then runs the Letter Wizard by using the **[RunLetterWizard](Word.Document.RunLetterWizard.md)** method.


```vb
Dim oLC as New LetterContent 
With oLC 
 .ReturnAddress = Application.UserAddress 
 .RecipientName = "Amy Anderson" 
 .RecipientAddress = "123 Main" & vbCr & "Bellevue, WA 98004" 
End With 
Documents.Add.RunLetterWizard LetterContent:=oLC, WizardMode:=True
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]