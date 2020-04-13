---
title: LetterContent.MailingInstructions property (Word)
keywords: vbawd10.chm161546354
f1_keywords:
- vbawd10.chm161546354
ms.prod: word
api_name:
- Word.LetterContent.MailingInstructions
ms.assetid: a31f4a82-984d-8aae-294d-9ffaaa889028
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.MailingInstructions property (Word)

Returns or sets the mailing instruction text for a letter created by the Letter Wizard (for example, "Certified Mail"). Read/write  **String**.


## Syntax

_expression_. `MailingInstructions`

 _expression_ An expression that returns a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example retrieves the Letter Wizard elements from the active document, changes the text of the mailing instructions, and then uses the **[SetLetterContent](Word.Document.SetLetterContent.md)** method to update the active document to reflect the changes.


```vb
Set myLetterContent = ActiveDocument.GetLetterContent 
myLetterContent.MailingInstructions = "Air Mail" 
ActiveDocument.SetLetterContent LetterContent:=myLetterContent
```

This example creates a new **LetterContent** object, sets several properties (including the mailing instruction text), and then runs the Letter Wizard by using the **[RunLetterWizard](Word.Document.RunLetterWizard.md)** method.




```vb
Set myContent = New LetterContent 
With myContent 
 .RecipientReference = "In reply to:" 
 .Salutation = "Hello" 
 .MailingInstructions = "Certified Mail" 
End With 
Documents.Add.RunLetterWizard LetterContent:=myContent
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]