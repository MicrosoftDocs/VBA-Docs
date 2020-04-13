---
title: EndnoteOptions object (Word)
ms.prod: word
api_name:
- Word.EndnoteOptions
ms.assetid: b63cf439-2297-fec9-ba36-66ad3f43dcbc
ms.date: 06/08/2017
localization_priority: Normal
---


# EndnoteOptions object (Word)

Represents the properties assigned to a range or selection of endnotes in a document.


## Remarks

Use the **EndnoteOptions** property of the **[Range](Word.Range.md)** or **[Selection](Word.Selection.md)** object to return an **EndnoteOptions** object.

Using the **EndnoteOptions** object, you can assign different endnote properties to different areas of a document. For example, you may want endnotes in the introduction of a long document to be displayed as lowercase Roman numerals, while in the rest of your document they are displayed as Arabic numerals. The following example uses the **[NumberingRule](Word.EndnoteOptions.NumberingRule.md)**, **[NumberStyle](Word.EndnoteOptions.NumberStyle.md)**, and **[StartingNumber](Word.EndnoteOptions.StartingNumber.md)** properties to format the endnotes in the first section ofthe active document.




```vb
Sub BookIntro() 
 Dim rngIntro As Range 
 
 'Sets the range as section one of the active document 
 Set rngIntro = ActiveDocument.Sections(1).Range 
 
 'Formats the EndnoteOptions properties 
 With rngIntro.EndnoteOptions 
 .NumberingRule = wdRestartSection 
 .NumberStyle = wdNoteNumberStyleLowercaseRoman 
 .StartingNumber = 1 
 End With 
End Sub
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]