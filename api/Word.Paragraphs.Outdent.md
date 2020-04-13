---
title: Paragraphs.Outdent method (Word)
keywords: vbawd10.chm156762446
f1_keywords:
- vbawd10.chm156762446
ms.prod: word
api_name:
- Word.Paragraphs.Outdent
ms.assetid: 94eda3f5-a67d-1e25-9851-65f64be5f472
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.Outdent method (Word)

Removes one level of indent for one or more paragraphs.


## Syntax

_expression_. `Outdent`

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Remarks

This method is equivalent to clicking the **Decrease Indent** button on the **Formatting** toolbar.


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


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]