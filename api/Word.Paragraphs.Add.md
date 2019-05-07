---
title: Paragraphs.Add method (Word)
keywords: vbawd10.chm156762117
f1_keywords:
- vbawd10.chm156762117
ms.prod: word
api_name:
- Word.Paragraphs.Add
ms.assetid: a75b7e4c-0a94-2bea-27bc-e6ad68ac075e
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.Add method (Word)

Returns a  **Paragraph** object that represents a new, blank paragraph added to a document.


## Syntax

_expression_.**Add** (_Range_)

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Optional| **Variant**|The range before which you want the new paragraph to be added. The new paragraph doesn't replace the range.|

## Return value

Paragraph


## Remarks

If Range isn't specified, the new paragraph is added after the selection or range or at the end of the document, depending on expression.


## Example

This example adds a paragraph after the selection.


```vb
Selection.Paragraphs.Add
```

This example adds a paragraph mark before the first paragraph in the selection.




```vb
Selection.Paragraphs.Add Range:=Selection.Paragraphs(1).Range
```

This example adds a paragraph mark before the second paragraph in the active document.




```vb
ActiveDocument.Paragraphs.Add _ 
 Range:=ActiveDocument.Paragraphs(2).Range
```

This example adds a new paragraph mark at the end of the active document.




```vb
ActiveDocument.Paragraphs.Add
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
