---
title: View.Type property (Project)
ms.prod: project-server
api_name:
- Project.View.Type
ms.assetid: ba42ed15-75ba-fad6-588a-3c4b8f42bad5
ms.date: 06/08/2017
localization_priority: Normal
---


# View.Type property (Project)

Gets the type of item in the view, such as tasks or resources. Read-only  **PjItemType**.


## Syntax

_expression_.**Type**

_expression_ A variable that represents a [View](./Project.View.md) object.


## Remarks

The **Type** property for a view can be one of the following **[PjItemType](Project.PjItemType.md)** constants: **pjOtherItem**, **pjResourceItem**, or **pjTaskItem**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]