---
title: LetterContent.EnclosureNumber property (Word)
keywords: vbawd10.chm161546357
f1_keywords:
- vbawd10.chm161546357
ms.prod: word
api_name:
- Word.LetterContent.EnclosureNumber
ms.assetid: e4bc5df9-a59a-562b-758e-4774eb4dbb9e
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.EnclosureNumber property (Word)

Returns or sets the number of enclosures for a letter created by the Letter Wizard. Read/write  **String**.


## Syntax

_expression_. `EnclosureNumber`

_expression_ A variable that represents a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example displays the number of enclosures specified in the active document.


```vb
MsgBox ActiveDocument.GetLetterContent.EnclosureNumber
```

This example retrieves letter elements from the active document, changes the number of enclosures by setting the **[EnclosureNumber](Word.LetterContent.EnclosureNumber.md)** property, and then uses the **[SetLetterContent](Word.Document.SetLetterContent.md)** method to update the active document to reflect the changes.




```vb
Dim lcTemp As LetterContent 
 
Set lcTemp = ActiveDocument.GetLetterContent 
lcTemp.EnclosureNumber = "5" 
ActiveDocument.SetLetterContent LetterContent:=lcTemp
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]