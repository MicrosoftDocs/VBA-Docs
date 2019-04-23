---
title: Paragraph.Previous method (Word)
keywords: vbawd10.chm156696901
f1_keywords:
- vbawd10.chm156696901
ms.prod: word
api_name:
- Word.Paragraph.Previous
ms.assetid: 0ccc928e-26c3-d5e6-ea99-a3d9776fbdd1
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.Previous method (Word)

Returns the previous paragraph as a  **Paragraph** object.


## Syntax

_expression_.**Previous** (_Count_)

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Count_|Optional| **Variant**|The number of paragraphs by which you want to move back. The default value is 1.|

## Return value

Paragraph


## Example

This example selects the paragraph that precedes the selection in the active document.


```vb
Selection.Previous(Unit:=wdParagraph, Count:=1).Select
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]