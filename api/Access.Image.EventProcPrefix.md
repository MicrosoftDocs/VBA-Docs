---
title: Image.EventProcPrefix property (Access)
keywords: vbaac10.chm10363
f1_keywords:
- vbaac10.chm10363
ms.prod: access
api_name:
- Access.Image.EventProcPrefix
ms.assetid: 57817dd3-62ed-5595-8196-f914f1fda037
ms.date: 02/21/2019
localization_priority: Normal
---


# Image.EventProcPrefix property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write **String**.


## Syntax

_expression_.**EventProcPrefix**

_expression_ A variable that represents an **[Image](Access.Image.md)** object.


## Remarks

For example, if you have a command button with an event procedure named **Details_Click**, the **EventProcPrefix** property returns the string **Details**.

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character ( _ ).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]