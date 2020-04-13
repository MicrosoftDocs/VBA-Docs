---
title: Document.FormattingShowNumbering property (Word)
keywords: vbawd10.chm158007747
f1_keywords:
- vbawd10.chm158007747
ms.prod: word
api_name:
- Word.Document.FormattingShowNumbering
ms.assetid: 2f0d8c8c-64a0-7939-e4be-99ed58ed696f
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.FormattingShowNumbering property (Word)

 **True** for Microsoft Word to display number formatting in the **Styles and Formatting** task pane. Read/write **Boolean**.


## Syntax

_expression_. `FormattingShowNumbering`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example enables displaying number formatting in the **Styles and Formatting** task pane.


```vb
Sub ShowClearFormatting() 
 With ActiveDocument 
 .FormattingShowClear = False 
 .FormattingShowFilter = wdShowFilterFormattingInUse 
 .FormattingShowFont = True 
 .FormattingShowNumbering = True 
 .FormattingShowParagraph = True 
 End With 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]