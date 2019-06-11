---
title: ParagraphFormat.Duplicate method (Publisher)
keywords: vbapb10.chm5439510
f1_keywords:
- vbapb10.chm5439510
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.Duplicate
ms.assetid: 83156999-7867-05c2-9e85-4cc0f580ac6e
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.Duplicate method (Publisher)

Creates a duplicate of the specified **ParagraphFormat** object and then returns the new **ParagraphFormat** object.


## Syntax

_expression_.**Duplicate**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

ParagraphFormat


## Example

The following example duplicates the paragraph formatting information from the text range in shape one on page one of the active publication and applies it to the text range in shape two.

```vb
Dim pfTemp As ParagraphFormat 
 
With ActiveDocument.Pages(1) 
 Set pfTemp = .Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat.Duplicate 
 .Shapes(2).TextFrame _ 
 .TextRange.ParagraphFormat = pfTemp 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]