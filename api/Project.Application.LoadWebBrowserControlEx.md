---
title: Application.LoadWebBrowserControlEx method (Project)
keywords: vbapj.chm54
f1_keywords:
- vbapj.chm54
ms.prod: project-server
api_name:
- Project.Application.LoadWebBrowserControlEx
ms.assetid: 2dca75d3-30ad-ecd0-a465-1190234b9b9b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.LoadWebBrowserControlEx method (Project)

Displays HTML pages within Project when the  **Project Guide** is shown or hidden.


## Syntax

_expression_. `LoadWebBrowserControlEx`( `_TargetPage_`, `_WrapperPage_`, `_FunctionalityName_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TargetPage_|Required|**String**| A numeric ID that identifies the HTML target page that needs to be displayed. **TargetPage** can also be set to a URL, an XML stream, a pointer to an XML file, or any other string value.|
| _WrapperPage_|Optional|**Variant**|A pointer to an HTML page that provides wrapper functionality for the page being displayed in Project. The wrapper page contains event-handling code that allows Project functionality, such as saving files or changing views, to work when a webpage is being displayed. The WrapperPage parameter is used only when the  **Project Guide** is hidden. When the **Project Guide** is shown, mainpage.htm is used as the wrapper page, and a WrapperPage parameter, if specified, is ignored. If no WrapperPage parameter is specified, Project uses the default wrapper page, gbui://wrapper.htm.|
| _FunctionalityName_|Optional|**Variant**|Name of the Project Guide function in the goal area.|

## Return value

 **Boolean**


## Remarks

When the  **Project Guide** is hidden, the method loads the Web Browser Control within Project and issues the **LoadWebPage** event. When the **Project Guide** is shown, the method only issues the **LoadWebPage** event.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]