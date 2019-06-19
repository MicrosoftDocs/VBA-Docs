---
title: Comment.Stat property (Visio)
ms.prod: visio
ms.assetid: f457598c-af42-cb83-ecd2-4fd42898ea16
ms.date: 06/08/2017
localization_priority: Normal
---


# Comment.Stat property (Visio)

Returns status information for an object. Read-only  **Integer**.


## Syntax

_expression_.**Stat**

_expression_ A variable that represents a **[Comment](Visio.Comment.md)** object.


## Remarks

If an object is a reference to an entity in a document, and if that document closes, the  **Stat** property returns a value in which the **visStatClosed** bit is set.

If an object is a reference to an entity that has been deleted, the  **Stat** property returns a value in which the **visStatDeleted** bit is set.

A Component Object Model (COM) object, such as a Microsoft Visio  **Document** object, lives as long as it is held (pointed to) by a client, even if the object is logically in a deleted or closed state.


## Property value

 **INT16**


## See also


[Comment Object](Visio.comment.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]