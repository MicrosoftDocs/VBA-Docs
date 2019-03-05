---
title: WebBrowserControl.EventProcPrefix property (Access)
keywords: vbaac10.chm14683
f1_keywords:
- vbaac10.chm14683
ms.prod: access
api_name:
- Access.WebBrowserControl.EventProcPrefix
ms.assetid: 8dbf1fee-d9ab-ff0c-5571-e606c19fbf94
ms.date: 02/21/2019
localization_priority: Normal
---


# WebBrowserControl.EventProcPrefix property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write **String**.


## Syntax

_expression_.**EventProcPrefix**

_expression_ A variable that represents a **[WebBrowserControl](Access.WebBrowserControl.md)** object.


## Remarks

For example, if you have a command button with an event procedure named **Details_Click**, the **EventProcPrefix** property returns the string **Details**.

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character ( _ ).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]