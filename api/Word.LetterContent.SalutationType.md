---
title: LetterContent.SalutationType property (Word)
keywords: vbawd10.chm161546351
f1_keywords:
- vbawd10.chm161546351
ms.prod: word
api_name:
- Word.LetterContent.SalutationType
ms.assetid: f312bdfd-a10d-144d-4b99-0984707d13cb
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.SalutationType property (Word)

Returns or sets the type of salutation for a letter created by the Letter Wizard. Read/write  **WdSalutationType**.


## Syntax

_expression_. `SalutationType`

_expression_ Required. A variable that represents a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example creates a new **LetterContent** object, sets several properties (including the salutation text), and then runs the Letter Wizard by using the **RunLetterWizard** method.


```vb
Set myContent = New LetterContent 
myContent.SalutationType = wdSalutationBusiness 
Documents.Add.RunLetterWizard _ 
 LetterContent:=myContent, WizardMode:=True
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]