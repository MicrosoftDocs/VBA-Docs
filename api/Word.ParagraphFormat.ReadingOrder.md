---
title: ParagraphFormat.ReadingOrder property (Word)
keywords: vbawd10.chm156434563
f1_keywords:
- vbawd10.chm156434563
ms.prod: word
api_name:
- Word.ParagraphFormat.ReadingOrder
ms.assetid: 4a22e638-2af8-096a-d45c-2eed21dc8002
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.ReadingOrder property (Word)

Returns or sets the reading order of the specified paragraphs without changing their alignment. Read/write  **WdReadingOrder**.


## Syntax

_expression_.**ReadingOrder**

_expression_ Required. A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Remarks

Use the **[LtrPara](Word.Selection.LtrPara.md)**, **[LtrRun](Word.Selection.LtrRun.md)**, **[RtlPara](Word.Selection.RtlPara.md)**, and **[RtlRun](Word.Selection.RtlRun.md)** methods of the **Selection** object to change the paragraph alignment along with the reading order.


## Example

This example sets the reading order of the first paragraph to right-to-left.


```vb
ActiveDocument.Paragraphs(1).ReadingOrder = _ 
 wdReadingOrderRtl
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]