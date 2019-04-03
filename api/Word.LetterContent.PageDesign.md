---
title: LetterContent.PageDesign property (Word)
keywords: vbawd10.chm161546343
f1_keywords:
- vbawd10.chm161546343
ms.prod: word
api_name:
- Word.LetterContent.PageDesign
ms.assetid: 8544d8c1-3e43-22f5-794f-8bd7636f8a0e
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.PageDesign property (Word)

Returns or sets the name of the template attached to the document created by the Letter Wizard. Read/write  **String**.


## Syntax

_expression_. `PageDesign`

 _expression_ An expression that returns a '[LetterContent](Word.LetterContent.md)' object.


## Example

This example creates a new  **LetterContent** object, includes the header and footer from the Contemporary Letter template, and then runs the Letter Wizard by using the **[RunLetterWizard](Word.Document.RunLetterWizard.md)** method.


```vb
Set myContent = New LetterContent 
With myContent 
 .PageDesign = "C:\MSOffice\Templates\" _ 
 & "Letters & Faxes\Contemporary Letter.dot" 
 .IncludeHeaderFooter = True 
End With 
Documents.Add.RunLetterWizard LetterContent:=myContent
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]