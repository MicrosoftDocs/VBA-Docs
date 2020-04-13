---
title: LetterContent.LetterStyle property (Word)
keywords: vbawd10.chm161546344
f1_keywords:
- vbawd10.chm161546344
ms.prod: word
api_name:
- Word.LetterContent.LetterStyle
ms.assetid: fdb8e106-bb80-468d-4330-e601d3a52938
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.LetterStyle property (Word)

Returns or sets the layout of a letter created by the Letter Wizard. Read/write  **WdLetterStyle**.


## Syntax

_expression_. `LetterStyle`

_expression_ Required. A variable that represents a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example creates a new LetterContent object, selects a letter style, and then runs the Letter Wizard by using the **RunLetterWizard** method.


```vb
Set aLetterContent = New LetterContent 
aLetterContent.LetterStyle = wdFullBlock 
ActiveDocument.RunLetterWizard _ 
 LetterContent:=aLetterContent, WizardMode:=True
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]