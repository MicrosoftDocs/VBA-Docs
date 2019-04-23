---
title: LetterContent.Salutation property (Word)
keywords: vbawd10.chm161546350
f1_keywords:
- vbawd10.chm161546350
ms.prod: word
api_name:
- Word.LetterContent.Salutation
ms.assetid: 115a740f-720f-a7d7-df68-148cd36b22c0
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.Salutation property (Word)

Returns or sets the salutation text for a letter created by the Letter Wizard. Read/write  **String**.


## Syntax

_expression_. `Salutation`

 _expression_ An expression that returns a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example creates a new  **LetterContent** object, sets several properties (including the salutation text), and then runs the Letter Wizard by using the **[RunLetterWizard](Word.Document.RunLetterWizard.md)** method.


```vb
Set myContent = New LetterContent 
myContent.Salutation ="Hello," 
Documents.Add.RunLetterWizard LetterContent:=myContent
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]