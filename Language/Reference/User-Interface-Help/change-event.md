---
title: Change event
keywords: fm20.chm5224938
f1_keywords:
- fm20.chm5224938
ms.prod: office
api_name:
- Office.Change
ms.assetid: 4bf23772-5ae0-dc1d-1152-b7ea01f7e702
ms.date: 11/15/2018
localization_priority: Normal
---


# Change event

Occurs when the **Value** property changes.

## Syntax

**Private Sub**_object_ _**Change( )**

The **Change** event syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Settings

The Change event occurs when the setting of the **Value** property changes, regardless of whether the change results from execution of code or a user action in the interface.

Following are some examples of actions that change the **Value** property:

- Clicking a **[CheckBox](checkbox-control.md)**, **[OptionButton](optionbutton-control.md)**, or **[ToggleButton](togglebutton-control.md)**.
    
- Entering or selecting a new text value for a **[ComboBox](combobox-control.md)**, **[ListBox](listbox-control.md)**, or **[TextBox](textbox-control.md)**.
    
- Selecting a different tab on a **[TabStrip](tabstrip-control.md)**.
    
- Moving the scroll box in a **[ScrollBar](scrollbar-control.md)**.
    
- Clicking the up arrow or down arrow on a **[SpinButton](spinbutton-control.md)**.
    
- Selecting a different page on a **[MultiPage](multipage-control.md)**.

## Remarks

The Change event procedure can synchronize or coordinate data displayed among controls. For example, you can use the Change event procedure of a **ScrollBar** to update the contents of a **TextBox** that displays the value of the **ScrollBar**. Or you can use a Change event procedure to display data and formulas in a work area and results in another area.

> [!NOTE] 
> In some cases, the Click event may also occur when the **Value** property changes. However, using the Change event is the preferred technique for detecting a new value for a property.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]