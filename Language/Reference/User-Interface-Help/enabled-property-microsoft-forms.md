---
title: Enabled property (Microsoft Forms)
keywords: fm20.chm5225035
f1_keywords:
- fm20.chm5225035
ms.prod: office
ms.assetid: 7e0320e4-91fa-2d2d-c484-70e54831e33b
ms.date: 11/16/2018
localization_priority: Normal
---


# Enabled property (Microsoft Forms)

Specifies whether a control can receive the [focus](../../Glossary/vbe-glossary.md#focus) and respond to user-generated events.

## Syntax

_object_.**Enabled** [= _Boolean_ ]

The **Enabled** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the object can respond to user-generated events.|

## Settings

The settings for _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|The control can receive the focus and respond to user-generated events, and is accessible through code (default).|
|**False**|The user cannot interact with the control by using the mouse, keystrokes, accelerators, or hotkeys. The control is generally still accessible through code.|

## Remarks

Use the **Enabled** property to enable and disable controls. A disabled control appears dimmed, while an enabled control does not. Also, if a control displays a bitmap, the bitmap is dimmed whenever the control is dimmed. If **Enabled** is **False** for an **[Image](image-control.md)**, the control does not initiate events but does not appear dimmed.

The **Enabled** and **Locked** properties work together to achieve the following effects:

- If **Enabled** and **Locked** are both **True**, the control can receive focus and appears normally (not dimmed) in the form. The user can copy, but not edit, data in the control.
    
- If **Enabled** is **True** and **Locked** is **False**, the control can receive focus and appears normally in the form. The user can copy and edit data in the control.
    
- If **Enabled** is **False** and **Locked** is **True**, the control cannot receive focus and is dimmed in the form. The user can neither copy nor edit data in the control.
    
- If **Enabled** and **Locked** are both **False**, the control cannot receive focus and is dimmed in the form. The user can neither copy nor edit data in the control.
    
You can combine the settings of the **Enabled** and the **TabStop** properties to prevent the user from selecting a command button with TAB, while still allowing the user to click the button. Setting **TabStop** to **False** means that the command button won't appear in the [tab order](../../Glossary/vbe-glossary.md#tab-order). However, if **Enabled** is **True**, the user can still click the command button, as long as **TakeFocusOnClick** is set to **True**.

When the user tabs into an enabled **[MultiPage](multipage-control.md)** or **[TabStrip](tabstrip-control.md)**, the first page or tab in the control receives the focus. If the first page or tab of a **MultiPage** or **TabStrip** is disabled, the first enabled page or tab of that control receives the focus. If all pages or tabs of a **MultiPage** or **TabStrip** are disabled, the control is disabled and cannot receive the focus.

If a **[Frame](frame-control.md)** is disabled, all controls it contains are disabled.

Clicking a disabled **[ListBox](listbox-control.md)** does not initiate the Click event.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]