---
title: Document.FormattingShowFilter property (Word)
keywords: vbawd10.chm158007748
f1_keywords:
- vbawd10.chm158007748
ms.prod: word
api_name:
- Word.Document.FormattingShowFilter
ms.assetid: 41509d69-9cee-bf85-6530-c5603b9c9136
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.FormattingShowFilter property (Word)

Sets or returns a  **WdShowFilter** constant that represents the styles and formatting displayed in the **Styles and Formatting** task pane. Read/write **Boolean**.


## Syntax

 _expression_. `FormattingShowFilter`

 _expression_ Required. A variable that represents a '[Document](Word.Document.md)' object.


## Example

This example filters formatting to show in the Styles and Formatting task pane only the formatting in use in the active document.


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

