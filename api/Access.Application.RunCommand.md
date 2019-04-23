---
title: Application.RunCommand method (Access)
keywords: vbaac10.chm12568
f1_keywords:
- vbaac10.chm12568
ms.prod: access
api_name:
- Access.Application.RunCommand
ms.assetid: 2731352f-7f2d-db3a-314c-e8a789755dd5
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.RunCommand method (Access)

The **RunCommand** method runs a built-in command.


## Syntax

_expression_.**RunCommand** (_Command_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Command_|Required|**[AcCommand](Access.AcCommand.md)**|An **AcCommand** constant that specifies the command to run.|

## Remarks

Each menu and toolbar command in Microsoft Access has an associated constant that you can use with the **RunCommand** method to run that command from Visual Basic.

You can't use the **RunCommand** method to run a command on a custom menu or toolbar. You can only use it with built-in menus and toolbars.

The **RunCommand** method replaces the **DoMenuItem** method of the **DoCmd** object.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]