---
title: Application.SelectRowEnd method (Project)
keywords: vbapj.chm2044
f1_keywords:
- vbapj.chm2044
ms.prod: project-server
api_name:
- Project.Application.SelectRowEnd
ms.assetid: 4aa9b311-46d7-2424-e675-6be0c61248f3
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SelectRowEnd method (Project)

Selects the last cell in the row containing the active cell.


## Syntax

_expression_. `SelectRowEnd`( `_Extend_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Extend_|Optional|**Boolean**|**True** if the current selection is extended to the last cell. The default value is **False**.|

## Return value

 **Boolean**


## Remarks

The **SelectRowEnd** method is only available when the Gantt Chart, Task Sheet, Task Usage view, Resource Sheet, or Resource Usage view is the active view.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]