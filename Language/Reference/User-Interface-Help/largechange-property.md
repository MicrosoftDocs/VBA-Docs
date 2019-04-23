---
title: LargeChange property
keywords: fm20.chm5225049
f1_keywords:
- fm20.chm5225049
ms.prod: office
api_name:
- Office.LargeChange
ms.assetid: 61187f0d-4bba-d761-2bcb-400de7b7d42e
ms.date: 11/16/2018
localization_priority: Normal
---


# LargeChange property

Specifies the amount of movement that occurs when the user clicks between the scroll box and scroll arrow.

## Syntax

_object_.**LargeChange** [= _Long_ ]

The **LargeChange** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. An integer that specifies the amount of change to the **Value** property.|

## Remarks

The **LargeChange** property applies only to the **[ScrollBar](scrollbar-control.md)**. It does not apply to the scrollbars in other controls such as a **[TextBox](textbox-control.md)** or a drop-down **[ComboBox](combobox-control.md)**.

The value of **LargeChange** is the amount by which the scrollbar's **[Value](value-property-microsoft-forms.md)** property changes when the user clicks the area between the scroll box and scroll arrow. The direction of the movement is always toward the place where the user clicks. For example, in a horizontal **ScrollBar**, clicking to the left of the scroll box moves the scroll box to the left. In a vertical **ScrollBar**, clicking above the scroll box moves the scroll box up.

**LargeChange** does not have units. Any integer is a valid setting for **LargeChange**. The recommended range of values is from -32,767 to +32,767, and the value must be between the values of the **[Max and Min](max-min-properties.md)** properties of the **ScrollBar**.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]