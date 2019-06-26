---
title: Application.SidepaneToggle method (Project)
keywords: vbapj.chm52
f1_keywords:
- vbapj.chm52
ms.prod: project-server
api_name:
- Project.Application.SidepaneToggle
ms.assetid: 882c9bef-f150-7128-a506-388dbe39558d
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SidepaneToggle method (Project)

Triggers the  **[WindowSidepaneDisplayChange](Project.Application.WindowSidepaneDisplayChange.md)** event, which shows or hides the side pane of the **Project Guide**. Deprecated in Project.


## Syntax

_expression_. `SidepaneToggle`( `_Show_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|**True** if Project shows the side pane for the **Project Guide**.  **False** if Project hides the side pane for the **Project Guide**.|

## Return value

 **Boolean**


## Remarks

The  **SidepaneToggle** method is used to change the side pane display state; you cannot use this method to return the current display state of the side pane in the **Project Guide**.


> [!NOTE] 
> The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of the Project Guide for new development.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]