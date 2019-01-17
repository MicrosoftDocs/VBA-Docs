---
title: Application.DetailStylesRemove Method (Project)
keywords: vbapj.chm964
f1_keywords:
- vbapj.chm964
ms.prod: project-server
api_name:
- Project.Application.DetailStylesRemove
ms.assetid: 67be5a7d-f066-f22c-7df1-834caeb7b6e2
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DetailStylesRemove Method (Project)

Removes a timescale data field from a usage view.


## Syntax

 _expression_. `DetailStylesRemove`( ` _Item_` )

 _expression_ A variable that represents an [Application](./Project.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Optional|**Long**|The timescale data field to remove. The default value is  **pjWork**.If the active view is the Resource Usage view, can be one of the **[PjTimescaledData](Project.PjTimescaledData.md)** constants. If the active view is the Task Usage view, can be one of the **[PjTimescaledData](Project.PjTimescaledData.md)** constants.|

## Return value

 **Boolean**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]