---
title: Find.ParagraphFormat property (Word)
keywords: vbawd10.chm162529298
f1_keywords:
- vbawd10.chm162529298
ms.prod: word
api_name:
- Word.Find.ParagraphFormat
ms.assetid: ae8bbbaa-700d-7469-30e4-f412e4a32e76
ms.date: 06/08/2017
localization_priority: Normal
---


# Find.ParagraphFormat property (Word)

Returns or sets a  **[ParagraphFormat](Word.ParagraphFormat.md)** object that represents the paragraph settings for the specified find operation. Read/write.


## Syntax

_expression_. `ParagraphFormat`

_expression_ A variable that represents a '[Find](Word.Find.md)' object.


## Example

This example finds all double-spaced paragraphs in the active document and replaces the formatting with 1.5-line spacing.


```vb
With ActiveDocument.Content.Find 
 .ClearFormatting 
 .ParagraphFormat.Space2 
 .Replacement.ClearFormatting 
 .Replacement.ParagraphFormat.Space15 
 .Execute FindText:="", ReplaceWith:="", _ 
 Replace:=wdReplaceAll 
End With
```


## See also


[Find Object](Word.Find.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]