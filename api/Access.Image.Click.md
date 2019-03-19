---
title: Image.Click event (Access)
keywords: vbaac10.chm14166
f1_keywords:
- vbaac10.chm14166
ms.prod: access
api_name:
- Access.Image.Click
ms.assetid: 1bca7597-b536-908e-c3fd-25f9dd5e1ab8
ms.date: 02/12/2019
localization_priority: Normal
---


# Image.Click event (Access)

The **Click** event occurs when the user presses and then releases a mouse button over an object.

> [!NOTE] 
> The functionality for the **Image** object's **Click** and **DoubleClick** events has been deprecated. If you want an image with click/double-click events, use instead a **Button** control and associate an image with that control to provide better accessibility. **Button** controls are part of the Tab Order loop, but **Image** controls are not. Existing applications will not be affected by this change.

## Syntax

_expression_.**Click**

_expression_ A variable that represents an **[Image](Access.Image.md)** object.

## Remarks

This event applies to a control containing a hyperlink.
    
To run a macro or event procedure when this event occurs, set the **OnClick** property to the name of the macro or to [Event Procedure].

For a control, this event occurs when the user:

- Clicks a control with the left mouse button. Clicking a control with the right or middle mouse button does not trigger this event.
    
- Clicks a control containing hyperlink data with the left mouse button. Clicking a control with the right or middle mouse button does not trigger this event. When the user moves the mouse pointer over a control containing hyperlink data, the mouse pointer changes to a "hand" icon. When the user clicks the mouse button, the hyperlink is activated, and then the **Click** event occurs.
    
- Selects an item in a combo box or list box, either by pressing the arrow keys and then pressing the Enter key or by clicking the mouse button.
    
- Presses Spacebar when a command button, check box, option button, or toggle button has the focus.
    
- Presses the Enter key on a form that has a command button whose **Default** property is set to Yes.
    
- Presses the Esc key on a form that has a command button whose **Cancel** property is set to Yes.
    
- Presses a control's access key. For example, if a command button's **Caption** property is set to &Go, pressing Alt+G triggers the event.

Typically, you attach a **Click** event procedure or macro to a command button to carry out commands and command-like actions. For the other applicable controls, use this event to trigger actions in response to one of the occurrences discussed earlier in this topic.

You can use a CancelEvent action in a DblClick macro to cancel the second **Click** event. For more information, see the **[DblClick](access.image.dblclick.md)** event topic.

To distinguish between the left, right, and middle mouse buttons, use the **MouseDown** and **MouseUp** events.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]