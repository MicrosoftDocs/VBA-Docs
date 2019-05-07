---
title: ViewSingle.Type property (Project)
ms.prod: project-server
api_name:
- Project.ViewSingle.Type
ms.assetid: 58b21a88-c71d-9949-5ca2-a0511d24467e
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewSingle.Type property (Project)

Gets the type of item in the single view, such as tasks or resources. Read-only  **PjItemType**.


## Syntax

_expression_.**Type**

_expression_ A variable that represents a [ViewSingle](./Project.ViewSingle.md) object.


## Remarks

The  **Type** property for a view can be one of the following **[PjItemType](Project.PjItemType.md)** constants: **pjOtherItem**, **pjResourceItem**, or **pjTaskItem**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]