---
title: MenuSets.Parent property (Visio)
keywords: vis_sdr.chm13414040
f1_keywords:
- vis_sdr.chm13414040
ms.prod: visio
api_name:
- Visio.MenuSets.Parent
ms.assetid: e6d68f93-2dcf-e132-d3de-f21254a75fbc
ms.date: 06/08/2017
localization_priority: Normal
---


# MenuSets.Parent property (Visio)

Determines the parent of an object. Read-only.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a **[MenuSets](Visio.MenuSets.md)** object.


## Return value

UIObject


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

In general, an object's parent is the object that contains it. For example, the parent of a  **Menu** object is the **Menus** collection that contains the **Menu** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]