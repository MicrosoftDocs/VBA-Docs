---
title: Application.GetThemedColor method (Project)
keywords: vbapj.chm131095
f1_keywords:
- vbapj.chm131095
ms.prod: project-server
api_name:
- Project.Application.GetThemedColor
ms.assetid: d7d464cd-a6d0-72b9-33cd-d5d9e7f30b80
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GetThemedColor method (Project)

Returns the color of the specified theme element type in the Project Guide. Deprecated in Project.


## Syntax

_expression_. `GetThemedColor`( `_elementType_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _elementType_|Required|**Long**|A valid theme element type. Can be one of the constants in the  **[PjThemeElement](Project.PjThemeElement.md)** enumeration.|

## Return value

 **Long**


## Remarks


> [!NOTE] 
> The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of the Project Guide for new development.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]