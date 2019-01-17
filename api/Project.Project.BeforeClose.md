---
title: Project.BeforeClose Event (Project)
ms.prod: project-server
api_name:
- Project.Project.BeforeClose
ms.assetid: 53ee16f4-2a6f-a575-7feb-90d1b92b9b07
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.BeforeClose Event (Project)

Occurs before a project is closed.


## Syntax

 _expression_. `BeforeClose`( ` _pj_`, )

 _expression_ An expression that returns a [Project](./Project(enumerations).md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that will be closed.|

## Return value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]