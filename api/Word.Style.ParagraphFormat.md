---
title: Style.ParagraphFormat property (Word)
keywords: vbawd10.chm153878537
f1_keywords:
- vbawd10.chm153878537
ms.prod: word
api_name:
- Word.Style.ParagraphFormat
ms.assetid: 83f6d48e-e13f-d5ab-c18f-6345dd6f4e9c
ms.date: 06/08/2017
localization_priority: Normal
---


# Style.ParagraphFormat property (Word)

Returns or sets a  **[ParagraphFormat](Word.ParagraphFormat.md)** object that represents the paragraph settings for the specified style. Read/write.


## Syntax

_expression_. `ParagraphFormat`

_expression_ A variable that represents a '[Style](Word.Style.md)' object.


## Example

This example modifies the Heading 2 style for the active document. Paragraphs formatted with this style are indented to the first tab stop and double-spaced.


```vb
With ActiveDocument.Styles(wdStyleHeading2).ParagraphFormat 
 .TabIndent(1) 
 .Space2 
End With
```


## See also


[Style Object](Word.Style.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]