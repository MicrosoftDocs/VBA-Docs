---
title: Paragraph.Indent method (Word)
keywords: vbawd10.chm156696909
f1_keywords:
- vbawd10.chm156696909
ms.prod: word
api_name:
- Word.Paragraph.Indent
ms.assetid: 5fc23149-8011-d465-0a73-f1f6e88d5a1e
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.Indent method (Word)

Indents one or more paragraphs by one level.


## Syntax

_expression_. `Indent`

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

This method is equivalent to clicking the  **Increase Indent** button on the **Formatting** toolbar.


## Example

This example indents all the paragraphs in the active document twice, and then it removes one level of the indent for the first paragraph.


```vb
With ActiveDocument.Paragraphs 
 .Indent 
 .Indent 
End With 
ActiveDocument.Paragraphs(1).Outdent
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]