---
title: ParagraphFormat.LeftIndent property (Publisher)
keywords: vbapb10.chm5439494
f1_keywords:
- vbapb10.chm5439494
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.LeftIndent
ms.assetid: f9cc3a86-d382-92d7-ec24-d13fc5e3d844
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.LeftIndent property (Publisher)

Returns or sets a **Variant** that represents the left indent value (in [points](../language/glossary/vbe-glossary.md#point)) for the specified paragraphs. Read/write.


## Syntax

_expression_.**LeftIndent**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

Variant


## Example

This example indents the paragraph at the cursor position 0.5 inch. This example assumes that the cursor is in a text box.

```vb
Sub IndentParagraph() 
 Selection.TextRange.ParagraphFormat.LeftIndent = 36 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]