---
title: MenuSet.RowIndex property (Visio)
keywords: vis_sdr.chm13314255
f1_keywords:
- vis_sdr.chm13314255
ms.prod: visio
api_name:
- Visio.MenuSet.RowIndex
ms.assetid: 70cd9ace-8792-07e3-f7a7-fcb7b3987dbf
ms.date: 06/08/2017
localization_priority: Normal
---


# MenuSet.RowIndex property (Visio)

Gets or sets the docking order of a **MenuSet** object in relation to other items in the same docking area. Read/write.


## Syntax

_expression_.**RowIndex**

_expression_ A variable that represents a **[MenuSet](Visio.MenuSet.md)** object.


## Return value

Integer


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Objects that have lower numbers are docked first. Several items can share the same row index. If two or more items share the same row index, the item most recently assigned is displayed first in its group.

Constants that represent the first and last positions (see the following table) are declared by the Visio type library in member **[VisUIBarRow](Visio.visuibarrow.md)**.

|Constant|Value|
|:-----|:-----|
| **visBarRowFirst**|0|
| **visBarRowLast**|-1|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]