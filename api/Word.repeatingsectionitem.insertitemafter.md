---
title: RepeatingSectionItem.InsertItemAfter method (Word)
keywords: vbawd10.chm227999747
f1_keywords:
- vbawd10.chm227999747
ms.prod: word
ms.assetid: c2c0159a-e6a4-0f45-d512-1d3debd17ca2
ms.date: 06/08/2017
localization_priority: Normal
---


# RepeatingSectionItem.InsertItemAfter method (Word)

Adds a repeating section item after the specified item and returns the new item.


## Syntax

_expression_. `InsertItemAfter`

_expression_ A variable that represents a 'RepeatingSectionItem' object.


## Return value

 **REPEATINGSECTIONITEM**


## Remarks

You can call this method on repeating section item content controls only.

If the [ContentControl.AllowInsertDeleteSection](Word.contentcontrol.allowinsertdeletesection.md) property is set to **False**, this method will return an error.


## See also


[RepeatingSectionItem Object](Word.repeatingsectionitem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]