---
title: Attachment.BeforeUpdate property (Access)
keywords: vbaac10.chm13937
f1_keywords:
- vbaac10.chm13937
ms.prod: access
api_name:
- Access.Attachment.BeforeUpdate
ms.assetid: 44a17114-bbb6-8ec9-89b5-db09cf60de98
ms.date: 02/07/2019
localization_priority: Normal
---


# Attachment.BeforeUpdate property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the **[BeforeUpdate](access.attachment.beforeupdate-event.md)** event occurs. Read/write **String**.


## Syntax

_expression_.**BeforeUpdate**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **BeforeUpdate** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]