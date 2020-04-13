---
title: Broadcast.Resume method (Word)
keywords: vbawd10.chm36438119
f1_keywords:
- vbawd10.chm36438119
ms.prod: word
ms.assetid: 7808f9fa-c307-9381-9067-e37c249f3010
ms.date: 06/08/2017
localization_priority: Normal
---


# Broadcast.Resume method (Word)

Resumes the specified broadcast.


## Syntax

_expression_. `Resume`

_expression_ A variable that represents a **[Broadcast](Word.broadcast.md)** object.


## Return value

 **VOID**


## Remarks

The **Resume** method returns an error (#4700) if the document is DRM protected, is already being broadcast (#4698), is not being broadcast at all (#4702), or has conflicting edits (is in merge mode, #4701).


## See also


[Broadcast Object](Word.broadcast.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]