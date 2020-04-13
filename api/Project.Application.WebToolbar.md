---
title: Application.WebToolbar method (Project)
keywords: vbapj.chm1321
f1_keywords:
- vbapj.chm1321
ms.prod: project-server
api_name:
- Project.Application.WebToolbar
ms.assetid: ff0f557f-ec63-0acd-da89-bc06c857524d
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WebToolbar method (Project)

Shows or hides the Web toolbar. Obsolete in Project.


## Syntax

_expression_. `WebToolbar`( `_Show_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|**True** if the Web toolbar is shown. The default is to toggle the current setting.|

## Return value

 **Boolean**


## Remarks

Project does not use toolbars; the **WebToolbar** method has no effect.

You can create a custom group on a tab in the ribbon that includes commands for webpages. For example, open the **Project Options** dialog box, choose **Customize Ribbon**, and then create a new group in a tab. Add commands such as  **Back**,  **Forward**,  **Stop**,  **Refresh**,  **Start Page**,  **Search the Web**, and  **Open Hyperlink**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]