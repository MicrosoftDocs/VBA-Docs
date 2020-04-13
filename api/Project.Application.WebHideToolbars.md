---
title: Application.WebHideToolbars method (Project)
keywords: vbapj.chm1306
f1_keywords:
- vbapj.chm1306
ms.prod: project-server
api_name:
- Project.Application.WebHideToolbars
ms.assetid: c6e323c9-b1a4-79bb-d714-b7ddaebbf619
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WebHideToolbars method (Project)

Shows or hides all toolbars except the **Menu** and **Web** toolbars. Obsolete in Project.


## Syntax

_expression_. `WebHideToolbars`( `_Hide_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Hide_|Optional|**Boolean**|**True** if all toolbars except the **Menu** and **Web** toolbars are hidden. The default value is **True** if toolbars other than **Menu** and **Web** are displayed, and **False** if they are not.|

## Return value

 **Boolean**


### Remarks

Project does not use toolbars.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]