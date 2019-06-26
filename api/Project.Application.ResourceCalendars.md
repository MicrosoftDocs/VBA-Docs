---
title: Application.ResourceCalendars method (Project)
keywords: vbapj.chm605
f1_keywords:
- vbapj.chm605
ms.prod: project-server
api_name:
- Project.Application.ResourceCalendars
ms.assetid: 8c40cfad-ec40-43a4-5698-de5abaea7243
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ResourceCalendars method (Project)

Displays the  **Change Working Time** dialog box, which prompts the user to manage calendars.


## Syntax

_expression_. `ResourceCalendars`( `_Index_`, `_Locked_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**String**|The resource index number or resource name.|
| _Locked_|Optional|**Boolean**|**False** if the user can set working time for selected dates for a resource. **True** if the fields are locked for editing. The default value is **False**.|

## Return value

 **Boolean**


## Remarks

The  **ResourceCalendars** method returns a trappable error (error code 1101) when applied to material resources.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]