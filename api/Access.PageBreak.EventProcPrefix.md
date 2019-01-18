---
title: PageBreak.EventProcPrefix property (Access)
keywords: vbaac10.chm11671
f1_keywords:
- vbaac10.chm11671
ms.prod: access
api_name:
- Access.PageBreak.EventProcPrefix
ms.assetid: abb7dc97-7bc9-8ab3-95ed-3b39a731df30
ms.date: 06/08/2017
localization_priority: Normal
---


# PageBreak.EventProcPrefix property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

_expression_. `EventProcPrefix`

_expression_ A variable that represents a [PageBreak](Access.PageBreak.md) object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


[PageBreak Object](Access.PageBreak.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]