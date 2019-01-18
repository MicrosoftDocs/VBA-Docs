---
title: ConditionalStyle.ParagraphFormat property (Word)
keywords: vbawd10.chm91029513
f1_keywords:
- vbawd10.chm91029513
ms.prod: word
api_name:
- Word.ConditionalStyle.ParagraphFormat
ms.assetid: 189e11aa-1bbe-575d-b538-8e8d0c35eaa3
ms.date: 06/08/2017
localization_priority: Normal
---


# ConditionalStyle.ParagraphFormat property (Word)

Returns or sets a  **[ParagraphFormat](Word.ParagraphFormat.md)** object that represents the paragraph settings for the specified conditional style. Read/write.


## Syntax

 _expression_. `ParagraphFormat`

 _expression_ A variable that represents a '[ConditionalStyle](Word.ConditionalStyle.md)' object.


## Example

This example modifies the Heading 2 style for the active document. Paragraphs formatted with this style are indented to the first tab stop and double-spaced.


```vb
With ActiveDocument.Styles(wdStyleHeading2).ParagraphFormat 
 .TabIndent(1) 
 .Space2 
End With
```


## See also


[ConditionalStyle Object](Word.ConditionalStyle.md)

