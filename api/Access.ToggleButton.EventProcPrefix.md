---
title: ToggleButton.EventProcPrefix property (Access)
keywords: vbaac10.chm11697
f1_keywords:
- vbaac10.chm11697
ms.prod: access
api_name:
- Access.ToggleButton.EventProcPrefix
ms.assetid: 80a9cfe1-87c1-b95d-f9a7-6afeca7c4755
ms.date: 02/21/2019
localization_priority: Normal
---


# ToggleButton.EventProcPrefix property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write **String**.


## Syntax

_expression_.**EventProcPrefix**

_expression_ A variable that represents a **[ToggleButton](Access.ToggleButton.md)** object.


## Remarks

For example, if you have a command button with an event procedure named **Details_Click**, the **EventProcPrefix** property returns the string **Details**.

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character ( _ ).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]