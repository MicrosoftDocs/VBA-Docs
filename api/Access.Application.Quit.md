---
title: Application.Quit method (Access)
keywords: vbaac10.chm12507
f1_keywords:
- vbaac10.chm12507
ms.prod: access
api_name:
- Access.Application.Quit
ms.assetid: 075ad885-f25d-ea2d-bf74-8ec915265c63
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.Quit method (Access)

The **[Quit](Access.Application.Quit.md)** method quits Microsoft Access. You can select one of several options for saving a database object before quitting.


## Syntax

_expression_.**Quit** (_Option_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Options_|Optional|**[AcQuitOption](Access.AcQuitOption.md)**|An **AcQuitOption** constant that indicates the action to take when quitting Access. The default value is **acQuitSaveAll**.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]