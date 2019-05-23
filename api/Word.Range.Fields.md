---
title: Range.Fields property (Word)
keywords: vbawd10.chm157155392
f1_keywords:
- vbawd10.chm157155392
ms.prod: word
api_name:
- Word.Range.Fields
ms.assetid: 106c1cb4-0836-3ff3-3138-223356a4a42c
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Fields property (Word)

Returns a  **[Fields](Word.fields.md)** collection that represents all the fields in the range. Read-only.


## Syntax

_expression_. `Fields`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Example

This example removes all the fields from the main text story and the footer in the active document.


```vb
For Each aField in ActiveDocument.Fields 
 aField.Delete 
Next aField 
Set myRange = ActiveDocument.Sections(1).Footers _ 
 (wdHeaderFooterPrimary).Range 
For Each aField In myRange.Fields 
 aField.Delete 
Next aField
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]