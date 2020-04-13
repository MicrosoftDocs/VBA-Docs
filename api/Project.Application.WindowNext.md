---
title: Application.WindowNext method (Project)
keywords: vbapj.chm2005
f1_keywords:
- vbapj.chm2005
ms.prod: project-server
api_name:
- Project.Application.WindowNext
ms.assetid: 10b5306d-038a-1b1c-9dec-8dd9d8b05dc3
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowNext method (Project)

Activates the next window in the order in which windows were opened.


## Syntax

_expression_. `WindowNext`( `_NoWrap_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NoWrap_|Optional|**Boolean**|**True** if using **WindowNext** on the last opened window doesn't wrap around to the first opened window. The default value is **False**.|

## Return value

 **Boolean**


## Remarks

The window order is the order in which windows are opened. The drop-down window list in the **Window** group of the **View** tab in the Ribbon contains the alphabetically sorted list of open windows.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]