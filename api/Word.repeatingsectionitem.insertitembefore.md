---
title: RepeatingSectionItem.InsertItemBefore method (Word)
keywords: vbawd10.chm227999746
f1_keywords:
- vbawd10.chm227999746
ms.prod: word
ms.assetid: 9848e875-56bb-6a68-f397-1ce8b59331dd
ms.date: 06/08/2017
localization_priority: Normal
---


# RepeatingSectionItem.InsertItemBefore method (Word)

Adds a repeating section item before the specified item and returns the new item.


## Syntax

_expression_. `InsertItemBefore`

_expression_ A variable that represents a 'RepeatingSectionItem' object.


## Return value

 **REPEATINGSECTIONITEM**


### Remarks

You can call this method on repeating section item content controls only.

If the [ContentControl.AllowInsertDeleteSection](Word.contentcontrol.allowinsertdeletesection.md) property is set to **False**, this method will return an error.


## See also


[RepeatingSectionItem Object](Word.repeatingsectionitem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]