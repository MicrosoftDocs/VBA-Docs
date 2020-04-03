---
title: SplitParts.Add method (Project)
ms.prod: project-server
api_name:
- Project.SplitParts.Add
ms.assetid: 91f6a47e-fdd9-b826-8b2c-776406c2f276
ms.date: 06/08/2017
localization_priority: Normal
---


# SplitParts.Add method (Project)

Adds a  **SplitPart** object to a **SplitParts** collection.


## Syntax

_expression_.**Add** (_StartSplitPartOn_, _EndSplitPartOn_)

_expression_ A variable that represents a 'SplitParts' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _StartSplitPartOn_|Required|**Variant**|The start date of the task portion.|
| _EndSplitPartOn_|Required|**Variant**|The end date of the task portion. If EndSplitPartOn is on or before the date specified with StartSplitPartOn, the portion is not created.|

## Remarks

If creating a new task portion would overlap any other portions in the same task, the non-coincident portions are added to the existing portion.


## See also


[SplitParts Collection Object](Project.splitparts.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]