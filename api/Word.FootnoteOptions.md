---
title: FootnoteOptions object (Word)
keywords: vbawd10.chm2596
f1_keywords:
- vbawd10.chm2596
ms.prod: word
api_name:
- Word.FootnoteOptions
ms.assetid: 5fdeb6d6-ce33-44f5-62c1-743fc3770457
ms.date: 06/08/2017
localization_priority: Normal
---


# FootnoteOptions object (Word)

Represents the properties assigned to a range or selection of footnotes in a document.


## Remarks

Use the **Range** or **Selection** object to return a **FootnoteOptions** object. Using the **FootnoteOptions** object, you can assign different footnote properties to different areas of a document. For example, you may want footnotes in the introduction of a long document to be displayed as lowercase letters, while in the rest of your document they are displayed as asterisks. The following example uses the **NumberingRule**, **NumberStyle**, and **StartingNumber** properties to format the footnotes in the first section of the active document.


```vb
Sub BookIntro() 
 Dim rngIntro As Range 
 
 'Sets the range as section one of the active document 
 Set rngIntro = ActiveDocument.Sections(1).Range 
 
 'Formats the EndnoteOptions properties 
 With rngIntro.FootnoteOptions 
 .NumberingRule = wdRestartPage 
 .NumberStyle = wdNoteNumberStyleLowercaseLetter 
 .StartingNumber = 1 
 End With 
End Sub
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]