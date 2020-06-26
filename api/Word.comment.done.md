---
title: Comment.Done property (Word)
keywords: vbawd10.chm154993653
f1_keywords:
- vbawd10.chm154993653
ms.prod: word
ms.assetid: 60b655ec-e523-13c4-2d26-1b0863b55a24
ms.date: 06/25/2020
localization_priority: Normal
---


# Comment.Done property (Word)

Returns or sets a  **Boolean** whose value is **True** if the specified comment has been marked closed. Read/write.

> [!IMPORTANT]
> This property has changed. The `Comment.Done` property is still available, but when setting the **Done** flag for a single comment reply, there will be no visible effect in the Redesigned Comments experience. The command will apply the **Done** flag, so when a user opens the document in the previous commenting experience, the comment reply is displayed as resolved or unresolved.

## Syntax

_expression_. `Done`

_expression_ A variable that represents a [Comment](./Word.Comment.md) object.


## Property value

 **BOOL**


## See also


[Comment Object](Word.Comment.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
