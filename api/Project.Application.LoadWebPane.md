---
title: Application.LoadWebPane event (Project)
ms.prod: project-server
api_name:
- Project.Application.LoadWebPane
ms.assetid: b9fefabb-3d0b-9aa7-6d3b-b8fd8000571d
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.LoadWebPane event (Project)

Occurs when Project loads a Web pane for  **Task Drivers**,  **Deliverables**, or the **Project/Resource Import Wizard**.


## Syntax

_expression_. `LoadWebPane`( `_Window_`, `_TargetPage_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required|**Window**|The window from where the **LoadWebBrowserControl** method was called.|
| _TargetPage_|Required|**String**|The same TargetPage parameter that was used to call the **LoadWebBrowserControl** method.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]