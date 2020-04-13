---
title: Paragraphs.ReadingOrder property (Word)
keywords: vbawd10.chm156762243
f1_keywords:
- vbawd10.chm156762243
ms.prod: word
api_name:
- Word.Paragraphs.ReadingOrder
ms.assetid: 9f3fccf3-7474-231d-21c7-f719174d7c82
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.ReadingOrder property (Word)

Returns or sets the reading order of the specified paragraphs without changing their alignment. Read/write  **WdReadingOrder**.


## Syntax

_expression_.**ReadingOrder**

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Remarks

Use the **[LtrPara](Word.Selection.LtrPara.md)**, **[LtrRun](Word.Selection.LtrRun.md)**, **[RtlPara](Word.Selection.RtlPara.md)**, and **[RtlRun](Word.Selection.RtlRun.md)** methods of the **[Selection](Word.Selection.md)** object to change the paragraph alignment along with the reading order.


## Example

This example sets the reading order of all paragraphs in the active document to right-to-left.


```vb
ActiveDocument.Paragraphs.ReadingOrder = _ 
 wdReadingOrderRtl
```


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]