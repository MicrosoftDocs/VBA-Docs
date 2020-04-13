---
title: Broadcast.Resume method (PowerPoint)
keywords: vbapp10.chm732008
f1_keywords:
- vbapp10.chm732008
ms.assetid: d141edba-f466-2d40-b177-3d3c416098ab
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# Broadcast.Resume method (PowerPoint)

Resumes the specified broadcast.


## Syntax

_expression_. `Resume`

_expression_ A variable that represents a [Broadcast](PowerPoint.Broadcast.md) object.


## Return value

 **VOID**


## Remarks

The **Resume** method returns an error (#4700) if the document is DRM protected, is already being broadcast (#4698), is not being broadcast at all (#4702), or has conflicting edits (is in merge mode, #4701).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]