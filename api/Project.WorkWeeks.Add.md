---
title: WorkWeeks.Add method (Project)
ms.prod: project-server
api_name:
- Project.WorkWeeks.Add
ms.assetid: 46469e7b-8309-4e77-c89f-2115b9498c7a
ms.date: 06/08/2017
localization_priority: Normal
---


# WorkWeeks.Add method (Project)

Adds a **WorkWeek** object to a **WorkWeeks** collection.


## Syntax

_expression_.**Add** (_Start_, _Finish_, _Name_)

 _expression_ An expression that returns a 'WorkWeeks' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Required|**Variant**|The start date of a **WorkWeek** object.|
| _Finish_|Optional|**Variant**|The finish date of a **WorkWeek** object.|
| _Name_|Optional|**String**|The name of a **WorkWeek** object.|

## Return value

 **WorkWeek**


## See also


[WorkWeeks Collection Object](Project.workweeks.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]