---
title: LetterContent.Letterhead property (Word)
keywords: vbawd10.chm161546345
f1_keywords:
- vbawd10.chm161546345
ms.prod: word
api_name:
- Word.LetterContent.Letterhead
ms.assetid: afd847ed-46b2-2539-a4b4-550094974614
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.Letterhead property (Word)

 **True** if space is reserved for a preprinted letterhead in a letter created by the Letter Wizard. Read/write **Boolean**. The **[LetterheadSize](Word.LetterContent.LetterheadSize.md)** property controls the size of the reserved letterhead space.


## Syntax

_expression_. `Letterhead`

 _expression_ An expression that returns a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example creates a new **LetterContent** object, reserves an inch of space at the top of the page for a preprinted letterhead, and then runs the Letter Wizard by using the **[RunLetterWizard](Word.Document.RunLetterWizard.md)** method.


```vb
Dim lcNew As LetterContent 
 
Set lcNew = New LetterContent 
 
With lcNew 
 .Letterhead = True 
 .LetterheadLocation = wdLetterTop 
 .LetterheadSize = InchesToPoints(1) 
End With 
ActiveDocument.RunLetterWizard _ 
 LetterContent:=lcNew, WizardMode:=True
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]