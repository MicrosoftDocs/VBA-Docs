---
title: Document.OMathBreakSub property (Word)
keywords: vbawd10.chm158007825
f1_keywords:
- vbawd10.chm158007825
ms.prod: word
api_name:
- Word.Document.OMathBreakSub
ms.assetid: a361f255-1392-eddc-7771-98e9db7c291a
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.OMathBreakSub property (Word)

Returns or sets a  **[WdOMathBreakSub](Word.WdOMathBreakSub.md)** constant that represents how Microsoft Word handles a subtraction operator that falls before a line break. Read/write.


## Syntax

_expression_. `OMathBreakSub`

 _expression_ An expression that returns a [Document](./Word.Document.md) object.


## Remarks

This property is used only when the **[OMathBreakBin](Word.Document.OMathBreakBin.md)** property is set to **wdOMathBreakBinRepeat**. Subtraction sometimes receives special treatment when a line break falls on a subtraction operator and the document setting is to repeat the subtraction operator on the following line, because two negatives make a positive. Some writers choose to convert one of the minus signs into a plus sign, and some choose to keep the two negatives.


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]