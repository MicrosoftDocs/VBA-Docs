---
title: Comment.Replies property (PowerPoint)
keywords: vbapp10.chm642014
f1_keywords:
- vbapp10.chm642014
ms.assetid: 3af06afb-e507-bb3b-901b-30bf6bbfa0ef
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# Comment.Replies property (PowerPoint)

Returns a [Comments](PowerPoint.Comments.md) collection of **Comment** objects that are children of the specified comment. Read-only.


## Syntax

_expression_. `Replies`

_expression_ A variable that represents a [Comment](PowerPoint.Comment.md) object.


## Remarks

Calling the [Add](PowerPoint.Comments.Add.md) method on the returned collection of replies adds a new reply, unless the collection was accessed from a reply to a reply.


## Property value

 **COMMENTS**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]