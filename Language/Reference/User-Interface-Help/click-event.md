---
title: Click Event
keywords: fm20.chm2000070
f1_keywords:
- fm20.chm2000070
ms.prod: office
api_name:
- Office.Click
ms.assetid: 655b57b1-74fc-75e9-eb8d-debb83afaea9
ms.date: 06/08/2017
---


# Click Event



Occurs in one of two cases:


- The user clicks a control with the mouse.
    
- The user definitively selects a value for a control with more than one possible value.

## Syntax

For MultiPage, TabStrip **Private Sub**_object_ _**Click(**_index_**As Long)**
For all other controls **Private Sub**_object_ _**Click( )**
The  **Click** event syntax has these parts:


|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _index_|Required. The index of the page or tab in a  **[MultiPage](multipage-control.md)** or **[TabStrip](tabstrip-control.md)** associated with this event.|

## Remarks

Of the two cases where the Click event occurs, the first case applies to the  **[CommandButton](commandbutton-control.md)**, **[Frame](frame-control.md)**, **[Image](image-control.md)**, **[Label](label-control.md)**, **[ScrollBar](scrollbar-control.md)**, and **[SpinButton](spinbutton-control.md)**. The second case applies to the **[CheckBox](checkbox-control.md)**, **[ComboBox](combobox-control.md)**, **[ListBox](listbox-control.md)**, **[MultiPage](multipage-control.md)**, **[TabStrip](tabstrip-control.md)**, and **[ToggleButton](togglebutton-control.md)**. It also applies to an **[OptionButton](optionbutton-control.md)** when the value changes to **True**.
The following are examples of actions that initiate the Click event:


- Clicking a blank area of a form or a disabled control (other than a list box) on the form.
    
- Clicking a  **[CommandButton](commandbutton-control.md)**. If the command button doesn't already have the [focus](../../Glossary/vbe-glossary.md#focus), the Enter event occurs before the Click event.
    
- Pressing the SPACEBAR when a  **[CommandButton](commandbutton-control.md)** has the focus.
    
- Clicking a control.
    
- Pressing ENTER on a form that has a command button whose  **Default** property is set to **True**, as long as no other command button has the focus.
    
- Pressing ESC on a form that has a command button whose  **Cancel** property is set to **True**, as long as no other command button has the focus.
    
- Pressing a control's [accelerator key](../../Glossary/glossary-vba.md#accelerator-key).
    

When the Click event results from clicking a control, the sequence of events leading to the Click event is:


1. MouseDown
    
2. MouseUp
    
3. Click
    

For some controls, the Click event occurs when the  **Value** property changes. However, using the Change event is the preferred technique for detecting a new value for a property. The following are examples of actions that initiate the Click event due to assigning a new value to a control:


- Clicking a different page or tab in a  **[MultiPage](multipage-control.md)** or **[TabStrip](tabstrip-control.md)**. The **Value** property for these controls reflects the current **Page** or **Tab**. Clicking the current page or tab does not change the control's value and does not initiate the Click event.
    
- Clicking a  **[CheckBox](checkbox-control.md)** or **[ToggleButton](togglebutton-control.md)**, pressing the SPACEBAR when one of these controls has the focus, pressing the accelerator key for one of these controls, or changing the value of the control in code.
    
- Changing the value of an  **[OptionButton](optionbutton-control.md)** to **True**. Setting one **[OptionButton](optionbutton-control.md)** in a group to **True** sets all other buttons in the group to **False**, but the Click event occurs only for the button whose value changes to **True**.
    
- Selecting a value for a  **[ComboBox](combobox-control.md)** or **[ListBox](listbox-control.md)** so that it unquestionably matches an item in the control's drop-down list. For example, if a list is not sorted, the first match for characters typed in the edit region may not be the only match in the list, so choosing such a value does not initiate the Click event. In a sorted list, you can use entry-matching to ensure that a selected value is a unique match for text the user types.
    

The Click event is not initiated when  **Value** is set to **Null**.

 **Note**  Clicking changes the value of a control, thus it initiates the Click event. When you right-click, the value of the control does not change, so it does not initiate the Click event.


