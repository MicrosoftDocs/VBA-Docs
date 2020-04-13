---
title: ScrollBar.LargeChange Property (Outlook Forms Script)
keywords: olfm10.chm2001360
f1_keywords:
- olfm10.chm2001360
ms.prod: outlook
ms.assetid: 1236ef08-7788-a345-e2a6-a3c647fe2675
ms.date: 06/08/2017
localization_priority: Normal
---


# ScrollBar.LargeChange Property (Outlook Forms Script)

Returns or sets a **Long** that specifies the amount of movement that occurs when the user clicks between the scroll box and scroll arrow. Read/write.


## Syntax

_expression_.**LargeChange**

_expression_ A variable that represents a **ScrollBar** object.


## Remarks

The **LargeChange** property specifies the amount of change to the **[Value](Outlook.scrollbar.value.md)** property.

The **LargeChange** property applies only to the **[ScrollBar](Outlook.scrollbar.md)**. It does not apply to the scrollbars in other controls such as a **[TextBox](Outlook.textbox.md)** or a drop-down **[ComboBox](Outlook.combobox.md)**.

The value of  **LargeChange** is the amount by which the **ScrollBar** control's **Value** property changes when the user clicks the area between the scroll box and scroll arrow. The direction of the movement is always toward the place where the user clicks. For example, in a horizontal **ScrollBar**, clicking to the left of the scroll box moves the scroll box to the left. In a vertical  **ScrollBar**, clicking above the scroll box moves the scroll box up.

 **LargeChange** does not have units. Any integer is a valid setting for **LargeChange**. The recommended range of values is from -32,767 to +32,767, and the value must be between the values of the  **[Max](Outlook.scrollbar.max.md)** and **[Min](Outlook.scrollbar.min.md)** properties of the **ScrollBar**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]