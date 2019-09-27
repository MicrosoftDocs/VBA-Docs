---
title: CheckBox object (Outlook Forms Script)
keywords: olfm10.chm2000470
f1_keywords:
- olfm10.chm2000470
ms.prod: outlook
ms.assetid: 1834855b-f96c-aaa1-24ce-81d1e4e4e1db
ms.date: 09/26/2019
localization_priority: Normal
---


# CheckBox object (Outlook Forms Script)

Displays the selection state of an item.


## Remarks

Use a **CheckBox** to give the user a choice between two values such as **Yes/No**, **True/False**, or **On/Off**. When the user selects a **CheckBox**, it displays a special mark (such as an **X**) and its current setting is **Yes**, **True**, or **On**. If the user does not select the **CheckBox**, it is empty and its setting is **No**, **False**, or Off. Depending on the value of the **[TripleState](Outlook.checkbox.triplestate.md)** property, a **CheckBox** can also have a **Null** value.

If a **CheckBox** is bound to a data source, changing the setting changes the value of that source. A disabled **CheckBox** shows the current value, but is dimmed and does not allow changes to the value from the user interface.

You can also use check boxes inside a group box to select one or more of a group of related items. For example, you can create an order form that contains a list of available items, with a **CheckBox** preceding each item. The user can select a particular item or items by checking the corresponding **CheckBox**.

The default property of a **CheckBox** is the **[Value](Outlook.checkbox.value.md)** property.

The **[ListBox](Outlook.listbox.md)** also lets you put a check mark by selected options. Depending on your application, you can use the **ListBox** instead of using a group of **CheckBox** controls.

## Events

|Name|Description|
|:-----|:-----|
| [Click](Outlook.checkbox(event).md)|Occurs when the user clicks inside the control.|

## Properties

|Name|Description|
|:-----|:-----|
| [Accelerator](Outlook.checkbox.accelerator.md)|Returns or sets the accelerator key for a control. Read/write.|
| [Alignment](Outlook.checkbox.alignment.md)|Returns or sets an **Integer** that indicates the position of a control relative to its caption. Read/write.|
| [AutoSize](Outlook.checkbox.autosize.md)|Returns or sets a **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.|
| [BackColor](Outlook.checkbox.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [BackStyle](Outlook.checkbox.backstyle.md)|Returns or sets an **Integer** that specifies the background style for an object. Read/write.|
| [Caption](Outlook.checkbox.caption.md)|Returns or sets a **String** that appears on an object to identify or describe it. Read/write.|
| [Enabled](Outlook.checkbox.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [ForeColor](Outlook.checkbox.forecolor.md)|Returns or sets a **Long** that specifies the foreground color of an object. Read/write.|
| [GroupName](Outlook.checkbox.groupname.md)|Returns or sets a **String** that identifies a group of mutually exclusive **CheckBox** controls. Read/write.|
| [Locked](Outlook.checkbox.locked.md)|Returns or sets a **Boolean** that specifies whether a control can be edited. Read/write.|
| [MouseIcon](Outlook.checkbox.mouseicon.md)|Returns a **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](Outlook.checkbox.mousepointer.md)|Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [Picture](Outlook.checkbox.picture.md)|Returns a **String** that specifies the full path name of a bitmap to display on a control. Read-only.|
| [PicturePosition](Outlook.checkbox.pictureposition.md)|Returns or sets an **Integer** that specifies the location of the picture relative to its caption. Read/write.|
| [SpecialEffect](Outlook.checkbox.specialeffect.md)|Returns or sets an **Integer** that specifies the visual appearance of an object. Read/write.|
| [TextAlign](Outlook.checkbox.textalign.md)|Returns or sets an **Integer** that specifies how text is aligned in a control. Read/write.|
| [TripleState](Outlook.checkbox.triplestate.md)|Returns or sets a **Boolean** that determines whether a user can specify, from the user interface, the **Null** state for a **CheckBox**. Read/write.|
| [Value](Outlook.checkbox.value.md)|Returns or sets a **Variant** that specifies whether the check box is selected. Read/write.|
| [WordWrap](Outlook.checkbox.wordwrap.md)|Returns or sets a **Boolean** that specifies whether the contents of a control automatically wrap at the end of a line and the control expands to fit the text. Read/write.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

