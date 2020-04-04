---
title: Resources.Add method (Project)
ms.prod: project-server
api_name:
- Project.Resources.Add
ms.assetid: 4fb69f50-4ba6-89a4-f586-3df268ae7fd5
ms.date: 06/08/2017
localization_priority: Normal
---


# Resources.Add method (Project)

Adds a **Resource** object to a **Resources** collection.


## Syntax

_expression_.**Add** (_Name_, _Before_)

_expression_ A variable that represents a 'Resources' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the new resource. The default value is an empty string ("").|
| _Before_|Optional|**Long**|The position of the resource in its containing collection. The default value is the position of the last item in the collection.|

## Return value

 **Resource**


## See also


[Resources Collection Object](Project.resources(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]