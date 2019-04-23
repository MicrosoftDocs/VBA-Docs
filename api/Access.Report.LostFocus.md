---
title: Report.LostFocus event (Access)
keywords: vbaac10.chm13888
f1_keywords:
- vbaac10.chm13888
ms.prod: access
api_name:
- Access.Report.LostFocus
ms.assetid: 8b80c2bc-8be4-1842-4011-0e6475b3a865
ms.date: 02/13/2019
localization_priority: Normal
---


# Report.LostFocus event (Access)

The **LostFocus** event occurs when the specified report loses the focus.


## Syntax

_expression_.**LostFocus**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

To run a macro or event procedure when these events occur, set the **[OnLostFocus](access.report.onlostfocus.md)** property to the name of the macro or to [Event Procedure].

This event occurs when the focus moves in response to a user action, such as pressing the Tab key or clicking the object, or when you use the **SetFocus** method in Visual Basic or the SelectObject, GoToRecord, GoToControl, or GoToPage action in a macro.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]