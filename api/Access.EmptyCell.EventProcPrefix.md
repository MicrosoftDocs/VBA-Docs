---
title: EmptyCell.EventProcPrefix property (Access)
keywords: vbaac10.chm14302
f1_keywords:
- vbaac10.chm14302
ms.prod: access
api_name:
- Access.EmptyCell.EventProcPrefix
ms.assetid: b8efbef8-4eaa-abb7-19c9-311af8448821
ms.date: 02/21/2019
localization_priority: Normal
---


# EmptyCell.EventProcPrefix property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write **String**.


## Syntax

_expression_.**EventProcPrefix**

_expression_ A variable that represents an **[EmptyCell](Access.EmptyCell.md)** object.


## Remarks

For example, if you have a command button with an event procedure named **Details_Click**, the **EventProcPrefix** property returns the string **Details**.

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character ( _ ).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]