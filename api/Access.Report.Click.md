---
title: Report.Click event (Access)
keywords: vbaac10.chm13889
f1_keywords:
- vbaac10.chm13889
ms.prod: access
api_name:
- Access.Report.Click
ms.assetid: 37bd4936-2f66-b434-ae54-5f76dd943c4c
ms.date: 02/12/2019
localization_priority: Normal
---


# Report.Click event (Access)

The **Click** event occurs when the user presses and then releases a mouse button over a report.


## Syntax

_expression_.**Click**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

To run a macro or event procedure when this event occurs, set the **OnClick** property to the name of the macro or to [Event Procedure].

On a report, this event occurs when the user clicks a blank area of the report.

To distinguish between the left, right, and middle mouse buttons, use the **MouseDown** and **MouseUp** events.

## See also

- [Report.DblClick event](access.report.dblclick.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]