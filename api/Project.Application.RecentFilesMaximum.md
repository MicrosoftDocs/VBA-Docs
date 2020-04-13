---
title: Application.RecentFilesMaximum property (Project)
keywords: vbapj.chm132538
f1_keywords:
- vbapj.chm132538
ms.prod: project-server
api_name:
- Project.Application.RecentFilesMaximum
ms.assetid: 005c7c09-1fbf-b807-ebe6-601c55e56c97
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.RecentFilesMaximum property (Project)

Gets or sets the maximum number of recently used files to display in the **Recent Projects** pane of the Backstage view. Read/write **Long**.


## Syntax

_expression_. `RecentFilesMaximum`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Remarks

The value of the **RecentFilesMaximum** property can be 0 to 50.

Setting the **RecentFilesMaximum** property to 0 also sets the **DisplayRecentFiles** property to **False**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]