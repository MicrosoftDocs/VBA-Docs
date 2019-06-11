---
title: ParagraphFormat.RightIndent property (Publisher)
keywords: vbapb10.chm5439495
f1_keywords:
- vbapb10.chm5439495
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.RightIndent
ms.assetid: bc3102d3-afc5-3f19-b98a-7f816e374d1a
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.RightIndent property (Publisher)

Returns or sets a **Variant** that represents the right indent (in [points](../language/glossary/vbe-glossary.md#point)) for the specified paragraphs. Read/write.


## Syntax

_expression_.**RightIndent**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

Variant


## Example

This example sets the right indent for all paragraphs in the active document to one inch from the right margin. The **[InchesToPoints](Publisher.Application.InchesToPoints.md)** method is used to convert inches to points. This example assumes that there is at least one shape on the first page of the active publication.

```vb
Sub SetRightIndent() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Paragraphs(1).ParagraphFormat _ 
 .RightIndent = InchesToPoints(1) 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]