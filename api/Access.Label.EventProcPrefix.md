---
title: Label.EventProcPrefix property (Access)
keywords: vbaac10.chm10188
f1_keywords:
- vbaac10.chm10188
ms.prod: access
api_name:
- Access.Label.EventProcPrefix
ms.assetid: 089ac12e-6ad3-4c0f-1025-be4c21f036c6
ms.date: 02/21/2019
localization_priority: Normal
---


# Label.EventProcPrefix property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write **String**.


## Syntax

_expression_.**EventProcPrefix**

_expression_ A variable that represents a **[Label](Access.Label.md)** object.


## Remarks

For example, if you have a command button with an event procedure named **Details_Click**, the **EventProcPrefix** property returns the string **Details**.

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character ( _ ).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]