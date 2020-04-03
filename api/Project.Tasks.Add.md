---
title: Tasks.Add method (Project)
ms.prod: project-server
api_name:
- Project.Tasks.Add
ms.assetid: a6e2186b-610c-0888-a22a-8b7deba3f53f
ms.date: 06/08/2017
localization_priority: Normal
---


# Tasks.Add method (Project)

Adds a **Task** object to a **Tasks** collection.


## Syntax

_expression_.**Add** (_Name_, _Before_)

_expression_ A variable that represents a 'Tasks' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the new task. The default value is an empty string ("").|
| _Before_|Optional|**Long**|The position of the task in its containing collection. The default value is the position of the last item in the collection.|

## Return value

 **Task**


## See also


[Tasks Collection Object](Project.tasks(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]