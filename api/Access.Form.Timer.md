---
title: Form.Timer event (Access)
keywords: vbaac10.chm13659
f1_keywords:
- vbaac10.chm13659
ms.prod: access
api_name:
- Access.Form.Timer
ms.assetid: 395c62a1-5731-01b8-a4ea-852bfb30572f
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.Timer event (Access)

The **Timer** event occurs for a form at regular intervals as specified by the form's **[TimerInterval](Access.Form.TimerInterval.md)** property.


## Syntax

_expression_.**Timer**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

To run a macro or event procedure when this event occurs, set the **OnTimer** property to the name of the macro or to [Event Procedure].

By running a macro or event procedure when a **Timer** event occurs, you can control what Microsoft Access does at every timer interval. For example, you might want to requery underlying records or repaint the screen at specified intervals.

The **TimerInterval** property setting of the form specifies the interval, in milliseconds, between **Timer** events. The interval can be between 0 and 2,147,483,647 milliseconds. Setting the **TimerInterval** property to 0 prevents the **Timer** event from occurring.

## Example

The following example demonstrates a digital clock that you can display on a form. A label control displays the current time according to your computer's system clock. 

To try the example, add the following event procedure to a form that contains a label named **Clock**. Set the form's **TimerInterval** property to 1000 milliseconds to update the clock every second.

```vb
Private Sub Form_Timer() 
    Clock.Caption = Time        ' Update time display. 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
