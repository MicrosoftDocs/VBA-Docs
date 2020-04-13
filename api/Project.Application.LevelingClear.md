---
title: Application.LevelingClear method (Project)
keywords: vbapj.chm612
f1_keywords:
- vbapj.chm612
ms.prod: project-server
api_name:
- Project.Application.LevelingClear
ms.assetid: fdd537eb-f9c2-c8d9-ec26-0f4af9a63c33
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.LevelingClear method (Project)

Removes the effects of leveling.


## Syntax

_expression_. `LevelingClear`( `_All_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _All_|Optional|**Boolean**|**True** if delays are removed from all tasks. **False** if delays are removed from selected tasks only.|

## Return value

 **Boolean**


## Remarks

Using the **LevelingClear** method without specifying any arguments displays the **Clear Leveling** dialog box.

The **LevelingClear** method has no effect if a task has a priority of 1000 (do not level).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]