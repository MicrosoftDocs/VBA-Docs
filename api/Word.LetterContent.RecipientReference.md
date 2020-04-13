---
title: LetterContent.RecipientReference property (Word)
keywords: vbawd10.chm161546352
f1_keywords:
- vbawd10.chm161546352
ms.prod: word
api_name:
- Word.LetterContent.RecipientReference
ms.assetid: e792b88e-b1f7-4a03-a966-ed519891b46d
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.RecipientReference property (Word)

Returns or sets the reference line (for example, "In reply to:") for a letter created by the Letter Wizard. Read/write  **String**.


## Syntax

_expression_. `RecipientReference`

 _expression_ An expression that returns a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example creates a new **LetterContent** object, sets several properties (including the reference line), and then runs the Letter Wizard by using the **[RunLetterWizard](Word.Document.RunLetterWizard.md)** method.


```vb
Set myContent = New LetterContent 
With myContent 
 .RecipientReference = "In reply to:" 
 .Salutation ="Hello" 
 .MailingInstructions = "Certified Mail" 
End With 
Documents.Add.RunLetterWizard LetterContent:=myContent
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]