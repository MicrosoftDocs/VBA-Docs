---
title: Document.OMathBreakBin property (Word)
keywords: vbawd10.chm158007824
f1_keywords:
- vbawd10.chm158007824
ms.prod: word
api_name:
- Word.Document.OMathBreakBin
ms.assetid: 7ec16236-3597-232b-f640-2a9c5713865e
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.OMathBreakBin property (Word)

Returns or sets a  **[WdOMathBreakBin](Word.WdOMathBreakBin.md)** constant that represents where Microsoft Word places binary operators when equations span two or more lines. Read/write.


## Syntax

_expression_. `OMathBreakBin`

 _expression_ An expression that returns a [Document](./Word.Document.md) object.


## Remarks

When the equation breaks on a binary operator—for example, an addition, subtraction, or multiplication operator—there are three different placements of the operator: before the break, after the break, and repeated both before and after the break.

When this property is set to  **wdOMathBreakBinRepeat**, use the **[OMathBreakSub](Word.Document.OMathBreakSub.md)** property to specify how Word treats subtraction operators that appear before a line break.


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]