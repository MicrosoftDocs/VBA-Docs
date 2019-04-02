---
title: LetterContent.IncludeHeaderFooter property (Word)
keywords: vbawd10.chm161546342
f1_keywords:
- vbawd10.chm161546342
ms.prod: word
api_name:
- Word.LetterContent.IncludeHeaderFooter
ms.assetid: 365fe58d-ef60-436e-a942-d43f12bafee8
ms.date: 06/08/2017
localization_priority: Normal
---


# LetterContent.IncludeHeaderFooter property (Word)

 **True** if the header and footer from the page design template are included in a letter created by the Letter Wizard. Read/write **Boolean**. Use the **[PageDesign](Word.LetterContent.PageDesign.md)** property to set the name of the template attached to a document created by the Letter Wizard.


## Syntax

_expression_. `IncludeHeaderFooter`

 _expression_ An expression that returns '[LetterContent](Word.LetterContent.md)' object.


## Example

This example creates a new  **LetterContent** object, includes the header and footer from the Contemporary Letter template, and then runs the Letter Wizard by using the **[RunLetterWizard](Word.Document.RunLetterWizard.md)** method.


```vb
Dim lcNew As LetterContent 
 
Set lcNew = New LetterContent 
 
With lcNew 
 .PageDesign = "C:\Program Files\Microsoft Office\" _ 
 & "Templates\1033\Contemporary Letter.dot" 
 .IncludeHeaderFooter = True 
End With 
 
Documents.Add.RunLetterWizard LetterContent:=lcNew
```


## See also


[LetterContent Object](Word.LetterContent.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]