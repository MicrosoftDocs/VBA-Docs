---
title: DoCmd.RunCommand method (Access)
keywords: vbaac10.chm4174
f1_keywords:
- vbaac10.chm4174
ms.prod: access
api_name:
- Access.DoCmd.RunCommand
ms.assetid: 5d4a4a3c-cea0-7f2c-8af7-51b65f7bdcf8
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.RunCommand method (Access)

The **RunCommand** method runs a built-in command.


## Syntax

_expression_.**RunCommand** (_Command_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Command_|Required|**[AcCommand](Access.AcCommand.md)**|An **AcCommand** constant that specifies the command to run.|

## Remarks

Each menu and toolbar command in Microsoft Access has an associated constant that you can use with the **RunCommand** method to run that command from Visual Basic.

You can't use the **RunCommand** method to run a command on a custom menu or toolbar. You can only use it with built-in menus and toolbars.

The **RunCommand** method replaces the **DoMenuItem** method of the **DoCmd** object.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
