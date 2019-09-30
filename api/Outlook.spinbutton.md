---
title: SpinButton object (Outlook Forms Script)
keywords: olfm10.chm2000630
f1_keywords:
- olfm10.chm2000630
ms.prod: outlook
ms.assetid: 3221b356-1e68-9e14-48ab-4a30c38aa685
ms.date: 06/08/2017
localization_priority: Normal
---


# SpinButton object (Outlook Forms Script)

Increments and decrements a value.


## Remarks

Clicking a **SpinButton** changes only the value of the **SpinButton**. You can write code that uses the **SpinButton** to update the displayed value of another control. For example, you can use a **SpinButton** to change the month, the day, or the year shown on a date. You can also use a **SpinButton** to scroll through a range of values or a list of items, or to change the value displayed in a text box.

To display a value updated by a **SpinButton**, you must assign the value of the **SpinButton** to the displayed portion of a control, such as the **[Caption](Outlook.label.caption.md)** property of a **[Label](Outlook.label.md)** or the **[Text](Outlook.textbox.text.md)** property of a **[TextBox](Outlook.textbox.md)**. To create a horizontal or vertical **SpinButton**, drag the sizing handles of the **SpinButton** horizontally or vertically on the form.

The default property for a **SpinButton** is the **[Value](Outlook.textbox.value.md)** property.


## Properties

|Name|Description|
|:-----|:-----|
| [BackColor](Outlook.spinbutton.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [Delay](Outlook.spinbutton.delay.md)|Returns or sets a **Long** that specifies the delay in milliseconds, between events on a [SpinButton](Outlook.spinbutton.md). Read/write.|
| [Enabled](Outlook.spinbutton.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [ForeColor](Outlook.spinbutton.forecolor.md)|Returns or sets a **Long** that specifies the foreground color of an object. Read/write.|
| [Max](Outlook.spinbutton.max.md)|Returns or sets a **Long** that specifies the maximum and minimum acceptable values for the **Value** property of a **SpinButton**. Read/write.|
| [Min](Outlook.spinbutton.min.md)|Returns or sets a **Long** that specifies the maximum and minimum acceptable values for the **Value** property of a **SpinButton**. Read/write.|
| [MouseIcon](Outlook.spinbutton.mouseicon.md)|Returns a **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](Outlook.spinbutton.mousepointer.md)|Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [Orientation](Outlook.spinbutton.orientation.md)|Returns or sets an **Integer** that specifies whether the control is oriented vertically or horizontally. Read/write.|
| [SmallChange](Outlook.spinbutton.smallchange.md)|Returns or sets an **Integer** that specifies the amount of movement that occurs when the user clicks either scroll arrow in a **SpinButton**. Read/write.|
| [Value](Outlook.spinbutton.value.md)|Returns or sets a **Variant** that specifies the state of a **SpinButton**. Read/write.|


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]