---
title: OptionButton object (Outlook Forms Script)
keywords: olfm10.chm2000580
f1_keywords:
- olfm10.chm2000580
ms.prod: outlook
ms.assetid: 8009dd64-44b5-3b66-e8d4-e3535e014396
ms.date: 06/08/2017
localization_priority: Normal
---


# OptionButton object (Outlook Forms Script)

Shows the selection status of one item in a group of choices.


## Remarks

Use an **OptionButton** to show whether a single item in a group is selected. Note that each **OptionButton** in a **[Frame](Outlook.frame.md)** is mutually exclusive.

If an **OptionButton** is bound to a data source, the **OptionButton** can show the value of that data source as either Yes/No, True/False, or On/Off. If the user selects the **OptionButton**, the current setting is Yes, True, or On. If the user does not select the **OptionButton**, the setting is No, False, or Off. For example, an **OptionButton** in an inventory-tracking application might show whether an item is discontinued. If the **OptionButton** is bound to a data source, then changing the setting changes the value of that data source. A disabled **OptionButton** is dimmed and does not show a value.

Depending on the value of the **[TripleState](Outlook.optionbutton.triplestate.md)** property, an **OptionButton** can also have a **Null** value.

You can also use an **OptionButton** inside a group box to select one or more of a group of related items. For example, you can create an order form with a list of available items, with an **OptionButton** preceding each item. The user can select a particular item by checking the corresponding **OptionButton** **OptionButton**.

The default property for an **OptionButton** is the **[Value](Outlook.optionbutton.value.md)** property.


## Events

|Name|Description|
|:-----|:-----|
| [Click](Outlook.optionbutton.click.md)|Occurs when the user definitively selects a value for the control that has more than one possible value, or when the value changes to **True**.|


## Properties

|Name|Description|
|:-----|:-----|
| [Accelerator](Outlook.optionbutton.accelerator.md)|Returns or sets the accelerator key for a control. Read/write.|
| [Alignment](Outlook.optionbutton.alignment.md)|Returns or sets an **Integer** that indicates the position of a control relative to its caption. Read/write.|
| [AutoSize](Outlook.optionbutton.autosize.md)|Returns or sets a **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.|
| [BackColor](Outlook.optionbutton.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [BackStyle](Outlook.optionbutton.backstyle.md)|Returns or sets an **Integer** that specifies the background style for an object. Read/write.|
| [Caption](Outlook.optionbutton.caption.md)|Returns or sets a **String** that appears on an object to identify or describe it. Read/write.|
| [Enabled](Outlook.optionbutton.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [ForeColor](Outlook.optionbutton.forecolor.md)|Returns or sets a **Long** that specifies the foreground color of an object. Read/write.|
| [GroupName](Outlook.optionbutton.groupname.md)|Returns or sets a **String** that identifies a group of mutually exclusive [OptionButton](Outlook.optionbutton.md) controls. Read/write.|
| [Locked](Outlook.optionbutton.locked.md)|Returns or sets a **Boolean** that specifies whether a control can be edited. Read/write.|
| [MouseIcon](Outlook.optionbutton.mouseicon.md)|Returns a **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](Outlook.optionbutton.mousepointer.md)|Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [Picture](Outlook.optionbutton.picture.md)|Returns a **String** that specifies the full path name of a bitmap to display on a control. Read-only.|
| [PicturePosition](Outlook.optionbutton.pictureposition.md)|Returns or sets an **Integer** that specifies the location of the picture relative to its caption. Read/write.|
| [SpecialEffect](Outlook.optionbutton.specialeffect.md)|Returns or sets an **Integer** that specifies the visual appearance of an object. Read/write.|
| [TextAlign](Outlook.optionbutton.textalign.md)|Returns or sets an **Integer** that specifies how text is aligned in a control. Read/write.|
| [TripleState](Outlook.optionbutton.triplestate.md)|Returns or sets a **Boolean** that determines whether the **OptionButton** supports the **Null** state. Read/write.|
| [Value](Outlook.optionbutton.value.md)|Returns or sets a **Variant** that specifies whether the option button is selected. Read/write.|
| [WordWrap](Outlook.optionbutton.wordwrap.md)|Returns or sets a **Boolean** that specifies whether the contents of a control automatically wrap at the end of a line and the control expands to fit the text. Read/write.|





[!include[Support and feedback](~/includes/feedback-boilerplate.md)]