---
title: Application.RemoveTimelineBar method (Project)
keywords: vbapj.chm158
f1_keywords:
- vbapj.chm158
ms.assetid: 8385d889-b81e-5422-a032-c7073fa7c65d
ms.date: 06/08/2017
ms.prod: project-server
localization_priority: Normal
---


# Application.RemoveTimelineBar method (Project)

Removes a  **Timeline** bar from the view. Introduced in Office 2016.


## Syntax

_expression_.**RemoveTimelineBar** ( _BarPosition_, _TimelineViewName_)

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _BarPosition_|Optional|**Integer**|Indicates the timeline bar to remove. If a number isn't specified, the selected bar is removed if applicable. The top bar is 0 and the next is 1, and so on. If a number is not specified, the selected bar is removed if one is selected. The last timeline bar cannot be removed.|
| _TimelineViewName_|Optional|**String**|Specifies the name of a timeline. The name can be the built-in timeline or an existing custom timeline such as "My Timeline". The default value is the name of the active timeline.|

## Return value

 **BOOLEAN**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]