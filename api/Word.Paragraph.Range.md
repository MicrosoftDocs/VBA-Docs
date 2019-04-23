---
title: Paragraph.Range property (Word)
keywords: vbawd10.chm156696576
f1_keywords:
- vbawd10.chm156696576
ms.prod: word
api_name:
- Word.Paragraph.Range
ms.assetid: 6da6e452-b938-9e02-3d22-6f0cb0544b82
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.Range property (Word)

Returns a  **Range** object that represents the portion of a document that is contained within the specified paragraph.


## Syntax

_expression_.**Range**

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example applies the Heading 1 style to the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).Range.Style = wdStyleHeading1
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]