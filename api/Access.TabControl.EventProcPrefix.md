---
title: TabControl.EventProcPrefix property (Access)
keywords: vbaac10.chm12072
f1_keywords:
- vbaac10.chm12072
ms.prod: access
api_name:
- Access.TabControl.EventProcPrefix
ms.assetid: 86c32c0c-7132-9658-411f-4a0ad91ed7ff
ms.date: 02/21/2019
localization_priority: Normal
---


# TabControl.EventProcPrefix property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write **String**.


## Syntax

_expression_.**EventProcPrefix**

_expression_ A variable that represents a **[TabControl](Access.TabControl.md)** object.


## Remarks

For example, if you have a command button with an event procedure named **Details_Click**, the **EventProcPrefix** property returns the string **Details**.

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character ( _ ).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]