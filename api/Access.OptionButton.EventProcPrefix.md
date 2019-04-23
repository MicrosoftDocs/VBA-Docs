---
title: OptionButton.EventProcPrefix property (Access)
keywords: vbaac10.chm10565
f1_keywords:
- vbaac10.chm10565
ms.prod: access
api_name:
- Access.OptionButton.EventProcPrefix
ms.assetid: 95896310-8723-de8f-dec9-51fded5227bb
ms.date: 02/21/2019
localization_priority: Normal
---


# OptionButton.EventProcPrefix property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write **String**.


## Syntax

_expression_.**EventProcPrefix**

_expression_ A variable that represents an **[OptionButton](Access.OptionButton.md)** object.


## Remarks

For example, if you have a command button with an event procedure named **Details_Click**, the **EventProcPrefix** property returns the string **Details**.

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character ( _ ).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]