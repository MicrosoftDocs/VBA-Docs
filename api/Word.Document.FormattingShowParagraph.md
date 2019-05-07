---
title: Document.FormattingShowParagraph property (Word)
keywords: vbawd10.chm158007746
f1_keywords:
- vbawd10.chm158007746
ms.prod: word
api_name:
- Word.Document.FormattingShowParagraph
ms.assetid: b2fc92be-02f5-1ed5-aa8a-76e4ed725b49
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.FormattingShowParagraph property (Word)

 **True** for Microsoft Word to display paragraph formatting in the **Styles and Formatting** task pane. Read/write **Boolean**.


## Syntax

_expression_. `FormattingShowParagraph`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example enables displaying paragraph formatting in the  **Styles and Formatting** task pane.


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