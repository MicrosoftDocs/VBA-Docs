---
title: Paragraph.NoLineNumber property (Word)
keywords: vbawd10.chm156696681
f1_keywords:
- vbawd10.chm156696681
ms.prod: word
api_name:
- Word.Paragraph.NoLineNumber
ms.assetid: f713018a-1024-25fd-7d25-07c278426ba3
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.NoLineNumber property (Word)

 **True** if line numbers are repressed for the specified paragraph. Read/write **Long**.


## Syntax

_expression_. `NoLineNumber`

_expression_ A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

This property can be  **True**, **False**, or **wdUndefined**. Use the **[LineNumbering](Word.PageSetup.LineNumbering.md)** property of the **[PageSetup](Word.PageSetup.md)** object to set line numbers.


## Example

This example enables line numbering for the active document. The starting number is set to 1, and the numbering is continuous throughout all sections in the document. Line numbering is then repressed for the second paragraph.


```vb
With ActiveDocument.PageSetup.LineNumbering 
 .Active = True 
 .StartingNumber = 1 
 .CountBy = 1 
 .RestartMode = wdRestartContinuous 
End With 
ActiveDocument.Paragraphs(2).NoLineNumber = True
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]