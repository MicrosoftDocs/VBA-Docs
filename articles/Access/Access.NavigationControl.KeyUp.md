---
title: NavigationControl.KeyUp Event (Access)
keywords: vbaac10.chm14208
f1_keywords:
- vbaac10.chm14208
ms.prod: access
api_name:
- Access.NavigationControl.KeyUp
ms.assetid: 35e7a26d-617c-9e51-c246-1830cd180420
ms.date: 06/08/2017
---


# NavigationControl.KeyUp Event (Access)

The  **KeyUp** event occurs when the user releases a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.


## Syntax

 _expression_. **KeyUp**( ** _KeyCode_**, ** _Shift_** )

 _expression_ A variable that represents a **NavigationControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required|**Integer**||
| _Shift_|Required|**Integer**||

### Return Value

nothing


## Remarks

To run a macro or event procedure when these events occur, set the  **OnKeyUp** property to the name of the macro or to [Event Procedure].

For both events, the object with the focus receives all keystrokes. A form can have the focus only if it has no controls or all its visible controls are disabled.

A form will also receive all keyboard events, even those that occur for controls, if you set the  **KeyPreview** property of the form to Yes. With this property setting, all keyboard events occur first for the form, and then for the control that has the focus. You can respond to specific keys pressed in the form, regardless of which control has the focus. For example, you may want the key combination CTRL+X to always perform the same action on a form.

If you press and hold down a key, the  **KeyDown** and **KeyPress** events alternate repeatedly ( **KeyDown**, **KeyPress**, **KeyDown**, **KeyPress**, and so on) until you release the key, then the **KeyUp** event occurs.

Although the  **KeyUp** event occurs when most keys are pressed, it is typically used to recognize or distinguish between:


- Extended character keys, such as function keys.
    
- Navigation keys, such as HOME, END, PAGE UP, PAGE DOWN, UP ARROW, DOWN ARROW, RIGHT ARROW, LEFT ARROW, and TAB.
    
- Combinations of keys and standard keyboard modifiers (SHIFT, CTRL, or ALT keys).
    
- The numeric keypad and keyboard number keys.
    
The  **KeyUp** event does not occur when you press:


- The ENTER key if the form has a command button for which the  **Default** property is set to Yes.
    
- The ESC key if the form has a command button for which the  **Cancel** property is set to Yes.
    
The  **KeyUp** event occurs after any event for a control caused by pressing or sending the key. If a keystroke causes the focus to move from one control to another control, the **KeyDown** event occurs for the first control, while the **KeyPress** and **KeyUp** events occur for the second control.

To find out the ANSI character corresponding to the key pressed, use the  **KeyPress** event.

If a modal dialog box is displayed as a result of pressing or sending a key, the  **KeyDown** and **KeyPress** events occur, but the **KeyUp** event doesn't occur.


## See also


#### Concepts


[NavigationControl Object](Access.NavigationControl.md)

