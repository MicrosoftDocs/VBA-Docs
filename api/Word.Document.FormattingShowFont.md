---
title: Document.FormattingShowFont property (Word)
keywords: vbawd10.chm158007744
f1_keywords:
- vbawd10.chm158007744
ms.prod: word
api_name:
- Word.Document.FormattingShowFont
ms.assetid: ea13daf7-6b62-ad27-bf87-21dd19e90878
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.FormattingShowFont property (Word)

 **True** for Microsoft Word to display font formatting in the **Styles and Formatting** task pane. Read/write **Boolean**.


## Syntax

_expression_. `FormattingShowFont`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example enables display of font formatting in the **Styles and Formatting** task pane.


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