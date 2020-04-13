---
title: Application.LoadWebPage event (Project)
ms.prod: project-server
api_name:
- Project.Application.LoadWebPage
ms.assetid: 393115c4-6245-3a1a-3c98-a5ddc1416aa0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.LoadWebPage event (Project)

Occurs after the **LoadWebBrowserControl** method is called. The method loads the web browser control inside Project, and then the event is fired.


## Syntax

_expression_. `LoadWebPage`( `_Window_`, `_TargetPage_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required|**Window**|The window from where the LoadWebBrowserControl method was called.|
| _TargetPage_|Required|**String**|The same TargetPage parameter that was used to call the LoadWebBrowserControl method.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]