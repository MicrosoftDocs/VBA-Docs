---
title: ParagraphFormat.SpaceAfter property (Publisher)
keywords: vbapb10.chm5439496
f1_keywords:
- vbapb10.chm5439496
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.SpaceAfter
ms.assetid: 52f65636-862d-442e-e66f-5ff5c79ee7b0
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.SpaceAfter property (Publisher)

Returns or sets a **Variant** that represents the amount of spacing (in [points](../language/glossary/vbe-glossary.md#point)) after one or more paragraphs. Read/write.


## Syntax

_expression_.**SpaceAfter**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

Variant


## Example

This example sets the spacing before and after the third paragraph in the first shape on the first page of the active publication to 6 points.

```vb
Sub SetSpacingBeforeAfterParagraph() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Paragraphs(3).ParagraphFormat 
 .SpaceBefore = 6 
 .SpaceAfter = 6 
 End With 
End Sub
```

<br/>

This example sets spacing before and after all paragraphs in the first shape on the first page of the active publication to 6 points.

```vb
Sub SetSpacingBeforeAfterAllParagraph() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat 
 .SpaceBefore = 12 
 .SpaceAfter = 6 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]