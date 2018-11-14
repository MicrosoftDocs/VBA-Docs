---
title: ScrollBar Control
keywords: fm20.chm5224985
f1_keywords:
- fm20.chm5224985
ms.prod: office
ms.assetid: 73b0b5af-dfca-2ebd-bb94-c4660c710bc9
ms.date: 06/08/2017
---


# ScrollBar Control



Returns or sets the value of another control based on the position of the scroll box.

## Remarks

A  **[ScrollBar](scrollbar-control.md)** is a stand-alone control you can place on a form. It is visually like the scroll bar you see in certain objects such as a **[ListBox](listbox-control.md)** or the drop-down portion of a **[ComboBox](combobox-control.md)**. However, unlike the scroll bars in these examples, the stand-alone **[ScrollBar](scrollbar-control.md)** is not an integral part of any other control.
To use the  **[ScrollBar](scrollbar-control.md)** to set or read the value of another control, you must write code for the **ScrollBar's** events and methods. For example, to use the **[ScrollBar](scrollbar-control.md)** to update the value of a **[TextBox](textbox-control.md)**, you can write code that reads the **Value** property of the **[ScrollBar](scrollbar-control.md)** and then sets the **Value** property of the **[TextBox](textbox-control.md)**.
The default property for a  **[ScrollBar](scrollbar-control.md)** is the **Value** property.
The default event for a  **[ScrollBar](scrollbar-control.md)** is the Change event.

 **Note**  To create a horizontal or vertical  **[ScrollBar](scrollbar-control.md)**, drag the sizing handles of the **[ScrollBar](scrollbar-control.md)** horizontally or vertically on the form.


## Related Topics

[ScrollBar Object](../../../api/Outlook.scrollbar.object.md)


