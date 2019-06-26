---
title: Menu.State property (Visio)
keywords: vis_sdr.chm13114425
f1_keywords:
- vis_sdr.chm13114425
ms.prod: visio
api_name:
- Visio.Menu.State
ms.assetid: c670b944-56fd-d3f4-24ce-c0a57e6352a1
ms.date: 06/08/2017
localization_priority: Normal
---


# Menu.State property (Visio)

Determines a menu's state, pressed or not pressed. Read/write.


## Syntax

_expression_.**State**

_expression_ A variable that represents a **[Menu](Visio.Menu.md)** object.


## Return value

Integer


## Remarks

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The **State** property can be one of the following constants declared by the Visio type library in **VisUIButtonState**.

|Constant|Value|Description|
|:-----|:-----|:-----|
| **visButtonUp**|0|Button is not pressed|
| **visButtonDown**|-1|Button is pressed|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]