---
title: Form.Close event (Access)
keywords: vbaac10.chm13645
f1_keywords:
- vbaac10.chm13645
ms.prod: access
api_name:
- Access.Form.Close
ms.assetid: e65fe7e0-efc1-dabc-4b2c-787af465ade0
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.Close event (Access)

The **Close** event occurs when a form is closed and removed from the screen.


## Syntax

_expression_.**Close**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Return value

Nothing


## Remarks

To run a macro or event procedure when this event occurs, set the **OnClose** property to the name of the macro or to [Event Procedure].

The **Close** event occurs after the **Unload** event, which is triggered after the form is closed but before it is removed from the screen.

When you close a form, the following events occur in this order:

> **Unload** → **Deactivate** → **Close**

When the **Close** event occurs, you can open another window or request the user's name to make a log entry indicating who used the form or report.

The **Unload** event can be canceled, but the **Close** event can't.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]