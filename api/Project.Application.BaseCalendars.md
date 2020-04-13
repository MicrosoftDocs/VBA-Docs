---
title: Application.BaseCalendars method (Project)
keywords: vbapj.chm604
f1_keywords:
- vbapj.chm604
ms.prod: project-server
api_name:
- Project.Application.BaseCalendars
ms.assetid: 5ae675d2-1be3-eb98-6c35-ff36c3fccf30
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BaseCalendars method (Project)

Displays the **Change Working Time** dialog box, which prompts the user to change calendar properties.


## Syntax

_expression_. `BaseCalendars`( `_Index_`, `_Locked_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**String**|The calendar index number or calendar name.|
| _Locked_|Optional|**Boolean**|**True** if Project disables the **New** and **Options** buttons in the **Change Working Time** dialog box.|

## Return value

 **Boolean**


## Remarks

The **BaseCalendars** method has the same effect as the **Change Working Time** command on the **PROJECT** tab of the ribbon.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]