---
title: LetterContent.DateFormat property (Word)
keywords: vbawd10.chm161546341
f1_keywords:
- vbawd10.chm161546341
ms.prod: word
api_name:
- Word.LetterContent.DateFormat
ms.assetid: 4d23139a-1691-4548-f395-e46aed0306a6
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.DateFormat property (Word)

Returns or sets the date for a letter created by the Letter Wizard. Read/write  **String**.


## Syntax

_expression_. `DateFormat`

_expression_ A variable that represents a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example displays the date from the letter that appears in the active document.


```vb
MsgBox ActiveDocument.GetLetterContent.DateFormat
```

This example creates a new **LetterContent** object, sets the date line to the current date, and then runs the Letter Wizard by using the **[RunLetterWizard](Word.Document.RunLetterWizard.md)** method.




```vb
Dim lcNew As LetterContent 
 
Set lcNew = New LetterContent 
lcNew.DateFormat = Date$ 
ActiveDocument.RunLetterWizard LetterContent:=lcNew
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]