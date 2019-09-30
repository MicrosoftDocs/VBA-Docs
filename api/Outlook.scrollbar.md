---
title: ScrollBar object (Outlook Forms Script)
keywords: olfm10.chm2000610
f1_keywords:
- olfm10.chm2000610
ms.prod: outlook
ms.assetid: 9e0a0f3d-fb04-2180-3beb-306b09c10c01
ms.date: 06/08/2017
localization_priority: Normal
---


# ScrollBar object (Outlook Forms Script)

Returns or sets the value of another control based on the position of the scroll box.


## Remarks

A **ScrollBar** is a stand-alone control you can place on a form. It is visually like the scroll bar you see in certain objects such as a **[ListBox](Outlook.listbox.md)** or the drop-down portion of a **[ComboBox](Outlook.combobox.md)**. However, unlike the scroll bars in these controls, the stand-alone **ScrollBar** is not an integral part of any other control.

To use the **ScrollBar** to set or read the value of another control, you must write code that uses the **ScrollBar** control's **[Value](Outlook.scrollbar.value.md)** property. For example, to use the **ScrollBar** to update the value of a **[TextBox](Outlook.textbox.md)**, you can write code that reads the **Value** property of the **ScrollBar** and then sets the **[Value](Outlook.scrollbar.value.md)** property of the **TextBox**.

The default property for a **ScrollBar** is the **Value** property.

To create a horizontal or vertical **ScrollBar**, drag the sizing handles of the **ScrollBar** horizontally or vertically on the form.


## Properties

|Name|Description|
|:-----|:-----|
| [BackColor](Outlook.scrollbar.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [Delay](Outlook.scrollbar.delay.md)|Returns or sets a **Long** that specifies the delay in milliseconds, between events on a [ScrollBar](Outlook.scrollbar.md). Read/write.|
| [Enabled](Outlook.scrollbar.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [ForeColor](Outlook.scrollbar.forecolor.md)|Returns or sets a **Long** that specifies the foreground color of an object. Read/write.|
| [LargeChange](Outlook.scrollbar.largechange.md)|Returns or sets a **Long** that specifies the amount of movement that occurs when the user clicks between the scroll box and scroll arrow. Read/write.|
| [Max](Outlook.scrollbar.max.md)|Returns or sets a **Long** that specifies the maximum and minimum acceptable values for the [Value](Outlook.scrollbar.value.md) property of a **ScrollBar**. Read/write.|
| [Min](Outlook.scrollbar.min.md)|Returns or sets a **Long** that specifies the maximum and minimum acceptable values for the **Value** property of a **ScrollBar**. Read/write.|
| [MouseIcon](Outlook.scrollbar.mouseicon.md)|Returns a **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](Outlook.scrollbar.mousepointer.md)|Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [Orientation](Outlook.scrollbar.orientation.md)|Returns or sets an **Integer** that specifies whether the control is oriented vertically or horizontally. Read/write.|
| [ProportionalThumb](Outlook.scrollbar.proportionalthumb.md)|Returns or sets a **Boolean** that specifies whether the size of the scroll box is proportional to the scrolling region or fixed. Read/write.|
| [SmallChange](Outlook.scrollbar.smallchange.md)|Returns or sets an **Integer** that specifies the amount of movement that occurs when the user clicks either scroll arrow in a **ScrollBar**. Read/write.|
| [Value](Outlook.scrollbar.value.md)|Returns or sets a **Variant** that specifies the state of a **ScrollBar**. Read/write.|


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]