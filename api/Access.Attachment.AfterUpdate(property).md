---
title: Attachment.AfterUpdate property (Access)
keywords: vbaac10.chm13938
f1_keywords:
- vbaac10.chm13938
ms.prod: access
api_name:
- Access.Attachment.AfterUpdate
ms.assetid: 556fc6d2-3936-5cc7-0c4f-03274f00cfc2
ms.date: 06/08/2017
localization_priority: Normal
---


# Attachment.AfterUpdate property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the  **AfterUpdate** event occurs. Read/write **String**.


## Syntax

_expression_. `AfterUpdate`

_expression_ An expression that returns an [Attachment](Access.Attachment.md) object.


## Remarks

Valid values for this property are " _macroname_" where  _macroname_ is the name of a macro; "[Event Procedure]" which indicates the event procedure associated with the **AfterUpdate** event for the specified object; or " **=** _functionname_ **()** " where _functionname_ is the name of a user-defined function.


## See also


[Attachment Object](Access.Attachment.md)

