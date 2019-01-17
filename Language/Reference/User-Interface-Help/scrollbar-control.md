---
title: ScrollBar control
keywords: fm20.chm5224985
f1_keywords:
- fm20.chm5224985
ms.prod: office
ms.assetid: 73b0b5af-dfca-2ebd-bb94-c4660c710bc9
ms.date: 11/15/2018
localization_priority: Normal
---


# ScrollBar control

Returns or sets the value of another control based on the position of the scroll box.

## Remarks

A **ScrollBar** is a stand-alone control you can place on a form. It is visually like the scroll bar you see in certain objects such as a **[ListBox](listbox-control.md)** or the drop-down portion of a **[ComboBox](combobox-control.md)**. However, unlike the scroll bars in these examples, the stand-alone **ScrollBar** is not an integral part of any other control.

To use the **ScrollBar** to set or read the value of another control, you must write code for the events and methods of the **ScrollBar**. For example, to use the **ScrollBar** to update the value of a **[TextBox](textbox-control.md)**, you can write code that reads the **Value** property of the **ScrollBar** and then sets the **Value** property of the **TextBox**.

The default property for a **ScrollBar** is the **Value** property. The default event for a **ScrollBar** is the Change event.

> [!NOTE] 
> To create a horizontal or vertical **ScrollBar**, drag the sizing handles of the **ScrollBar** horizontally or vertically on the form.

## See also

- [ScrollBar object](../../../api/Outlook.scrollbar.object.md)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]