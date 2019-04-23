---
title: DoCmd.Quit method (Access)
keywords: vbaac10.chm4167
f1_keywords:
- vbaac10.chm4167
ms.prod: access
api_name:
- Access.DoCmd.Quit
ms.assetid: 2644084a-fd24-6271-7679-46c5f1b206d5
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.Quit method (Access)

The **Quit** method quits Microsoft Access. You can select one of several options for saving a database object before quitting.


## Syntax

_expression_.**Quit** (_Options_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Options_|Optional|**[AcQuitOption](Access.AcQuitOption.md)**|An **AcQuitOption** constant that indicates the action to take when quitting Access. The default value is **acQuitSaveAll**.|

## Remarks

The **Quit** method of the **DoCmd** object was added to provide backward compatibility for running the Quit action in Visual Basic code in Microsoft Access 95. We recommend that you use the existing **Quit** method of the **[Application](Access.Application.md)** object instead.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
