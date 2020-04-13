---
title: LetterContent.LetterheadSize property (Word)
keywords: vbawd10.chm161546347
f1_keywords:
- vbawd10.chm161546347
ms.prod: word
api_name:
- Word.LetterContent.LetterheadSize
ms.assetid: 05cc8dc3-fd22-ae58-6457-2daf2e6875f4
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.LetterheadSize property (Word)

Returns or sets the amount of space (in points) to be reserved for a preprinted letterhead in a letter created by the Letter Wizard. Read/write  **Single**.


## Syntax

_expression_. `LetterheadSize`

 _expression_ An expression that returns a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example retrieves the Letter Wizard elements from the active document, changes the preprinted letterhead settings, and then uses the **[SetLetterContent](Word.Document.SetLetterContent.md)** method to update the active document to reflect the changes.


```vb
Set myLetterContent = ActiveDocument.GetLetterContent 
With myLetterContent 
 .Letterhead = True 
 .LetterheadLocation = wdLetterTop 
 .LetterheadSize = InchesToPoints(1.5) 
End With 
ActiveDocument.SetLetterContent LetterContent:=myLetterContent
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]