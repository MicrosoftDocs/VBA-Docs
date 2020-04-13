---
title: Application.FilterNew method (Project)
keywords: vbapj.chm504
f1_keywords:
- vbapj.chm504
ms.prod: project-server
api_name:
- Project.Application.FilterNew
ms.assetid: 9289cf4f-ce29-695d-baf8-08316ed1e31b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.FilterNew method (Project)

Shows the **Filter Definition** dialog box, where the user can create a filter for a task-based view, resource-based view, or the default view filter.


## Syntax

_expression_. `FilterNew`( `_FilterType_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FilterType_|Optional|**PjFilterViewType**|Specifies whether the filter is for task information or resource information. Can be one of the following constants of the **[PjFilterViewType](Project.PjFilterViewType.md)** enumeration: **pjFilterViewTypeResource**, **pjFilterViewTypeTask**, or **pjFilterViewTypeUseView**. The default value is **pjFilterViewTypeUseView**.|

## Return value

 **Boolean**


## Remarks

Running the **FilterNew** method with no arguments corresponds to the **New Filter** command in the **Filter** drop-down list on the **VIEW** tab of the ribbon. That command brings up the **Filter Definition** dialog box, where the **Field Name** drop-down list contains fields that apply to the current view.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]