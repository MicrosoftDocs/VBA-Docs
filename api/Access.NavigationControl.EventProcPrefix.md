---
title: NavigationControl.EventProcPrefix property (Access)
keywords: vbaac10.chm11040
f1_keywords:
- vbaac10.chm11040
ms.prod: access
api_name:
- Access.NavigationControl.EventProcPrefix
ms.assetid: d59c7baf-7614-821b-92ce-582d6f90441c
ms.date: 02/21/2019
localization_priority: Normal
---


# NavigationControl.EventProcPrefix property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write **String**.


## Syntax

_expression_.**EventProcPrefix**

_expression_ A variable that represents a **[NavigationControl](Access.NavigationControl.md)** object.


## Remarks

For example, if you have a command button with an event procedure named **Details_Click**, the **EventProcPrefix** property returns the string **Details**.

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character ( _ ).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]