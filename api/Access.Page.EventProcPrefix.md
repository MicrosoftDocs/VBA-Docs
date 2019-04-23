---
title: Page.EventProcPrefix property (Access)
keywords: vbaac10.chm12145
f1_keywords:
- vbaac10.chm12145
ms.prod: access
api_name:
- Access.Page.EventProcPrefix
ms.assetid: 935843c6-cc50-016d-5569-87263670af99
ms.date: 02/21/2019
localization_priority: Normal
---


# Page.EventProcPrefix property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write **String**.


## Syntax

_expression_.**EventProcPrefix**

_expression_ A variable that represents a **[Page](Access.Page.md)** object.


## Remarks

For example, if you have a command button with an event procedure named **Details_Click**, the **EventProcPrefix** property returns the string **Details**.

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character ( _ ).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]