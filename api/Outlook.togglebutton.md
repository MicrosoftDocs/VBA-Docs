---
title: ToggleButton object (Outlook Forms Script)
keywords: olfm10.chm2000680
f1_keywords:
- olfm10.chm2000680
ms.prod: outlook
ms.assetid: 01ce5640-9f19-3c0e-1aa4-96d87074bf8b
ms.date: 06/08/2017
localization_priority: Normal
---


# ToggleButton object (Outlook Forms Script)

Shows the selection state of an item.


## Remarks

Use a **ToggleButton** to show whether an item is selected. If a **ToggleButton** is bound to a data source, the **ToggleButton** shows the current value of that data source as either Yes/No, True/False, On/Off, or some other choice of two settings. If the user selects the **ToggleButton**, the current setting is Yes, True, or On. If the user does not select the **ToggleButton**, the setting is No, False, or Off. If the **ToggleButton** is bound to a data source, changing the setting changes the value of that data source. A disabled **ToggleButton** shows a value, but is dimmed and does not allow changes from the user interface.

You can also use a **ToggleButton** inside a **[Frame](Outlook.frame.md)** to select one or more of a group of related items. For example, you can create an order form with a list of available items, with a **ToggleButton** preceding each item. The user can select a particular item by selecting the appropriate **ToggleButton**.

The default property of a **ToggleButton** is the **[Value](Outlook.togglebutton.value.md)** property.

The only event for a **ToggleButton** is the **[Click](Outlook.togglebutton.click.md)** event.


## Events

|Name|Description|
|:-----|:-----|
| [Click](Outlook.togglebutton.click.md)|Occurs when the user definitively selects a value for the control that has more than one possible value.|


## Properties

|Name|Description|
|:-----|:-----|
| [Accelerator](Outlook.togglebutton.accelerator.md)|Returns or sets the accelerator key for a control. Read/write.|
| [Alignment](Outlook.togglebutton.alignment.md)|Returns or sets an **Integer** that indicates the position of a control relative to its caption. Read/write.|
| [AutoSize](Outlook.togglebutton.autosize.md)|Returns or sets a **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.|
| [BackColor](Outlook.togglebutton.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [BackStyle](Outlook.togglebutton.backstyle.md)|Returns or sets an **Integer** that specifies the background style for an object. Read/write.|
| [Caption](Outlook.togglebutton.caption.md)|Returns or sets a **String** that appears on an object to identify or describe it. Read/write.|
| [Enabled](Outlook.togglebutton.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [ForeColor](Outlook.togglebutton.forecolor.md)|Returns or sets a **Long** that specifies the foreground color of an object. Read/write.|
| [GroupName](Outlook.togglebutton.groupname.md)|Returns or sets a **String** that identifies a group of mutually exclusive [ToggleButton](Outlook.togglebutton.md) controls. Read/write.|
| [Locked](Outlook.togglebutton.locked.md)|Returns or sets a **Boolean** that specifies whether a control can be edited. Read/write.|
| [MouseIcon](Outlook.togglebutton.mouseicon.md)|Returns a **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](Outlook.togglebutton.mousepointer.md)|Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [Picture](Outlook.togglebutton.picture.md)|Returns a **String** that specifies the full path name of a bitmap to display on a control. Read-only.|
| [PicturePosition](Outlook.togglebutton.pictureposition.md)|Returns or sets an **Integer** that specifies the location of the picture relative to its caption. Read/write.|
| [SpecialEffect](Outlook.togglebutton.specialeffect.md)|Returns or sets an **Integer** that specifies the visual appearance of an object. Read/write.|
| [TextAlign](Outlook.togglebutton.textalign.md)|Returns or sets an **Integer** that specifies how text is aligned in a control. Read/write.|
| [TripleState](Outlook.togglebutton.triplestate.md)|Returns or sets a **Boolean** that determines whether a user can specify, from the user interface, the **Null** state for a **ToggleButton**. Read/write.|
| [Value](Outlook.togglebutton.value.md)|Returns or sets a **Variant** that specifies whether the toggle button is selected. Read/write.|
| [WordWrap](Outlook.togglebutton.wordwrap.md)|Returns or sets a **Boolean** that specifies whether the contents of a control automatically wrap at the end of a line and the control expands to fit the text. Read/write.|





[!include[Support and feedback](~/includes/feedback-boilerplate.md)]