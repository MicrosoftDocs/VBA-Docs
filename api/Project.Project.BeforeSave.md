---
title: Project.BeforeSave Event (Project)
ms.prod: project-server
api_name:
- Project.Project.BeforeSave
ms.assetid: 6947661e-f77c-b766-b926-fd37818019b7
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.BeforeSave Event (Project)

Occurs before a project is saved.


## Syntax

 _expression_. `BeforeSave`( ` _pj_`, )

 _expression_ An expression that returns a [Project](./Project(enumerations).md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that will be saved.|

## Return value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]