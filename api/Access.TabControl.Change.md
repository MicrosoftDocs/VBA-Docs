---
title: TabControl.Change event (Access)
keywords: vbaac10.chm14278
f1_keywords:
- vbaac10.chm14278
ms.prod: access
api_name:
- Access.TabControl.Change
ms.assetid: e57d4b0b-0f9e-28e7-c1e0-6a1582f1cb0f
ms.date: 02/13/2019
localization_priority: Normal
---


# TabControl.Change event (Access)

The **Change** event occurs when you move from one page to another page.


## Syntax

_expression_.**Change**

_expression_ A variable that represents a **[TabControl](Access.TabControl.md)** object.


## Remarks

> [!NOTE] 
> Setting the value of a control by using a macro or Visual Basic doesn't trigger this event for the control. You must type the data directly into the control, or set the control's **Text** property.

To run a macro or event procedure when this event occurs, set the **[OnChange](access.tabcontrol.onchange.md)** property to the name of the macro or to [Event Procedure].

By running a macro or event procedure when a **Change** event occurs, you can coordinate data display among controls. You can also display data or a formula in one control and the results in another control.

Changing the data in a text box or combo box by using the keyboard causes keyboard events to occur in addition to control events like the **Change** event. For example, if you move to a new record and type an ANSI character in a text box in the record, the following events occur in this order:

> **KeyDown** → **KeyPress** → **BeforeInsert** → **Change** → **KeyUp**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]


