---
title: CurTargetX property
keywords: fm20.chm5225027
f1_keywords:
- fm20.chm5225027
ms.prod: office
api_name:
- Office.CurTargetX
ms.assetid: b0365f58-22db-34d2-9751-6c9d36598e08
ms.date: 11/15/2018
localization_priority: Normal
---


# CurTargetX property

Retrieves the preferred horizontal position of the insertion point in a multiline **[TextBox](textbox-control.md)** or **[ComboBox](combobox-control.md)**.

## Syntax

_object_.**CurTargetX**

The **CurTargetX** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Return values

The **CurTargetX** property retrieves the preferred position, measured in himetric units. A himetric is 0.0001 meter.

## Remarks

The [target](../../Glossary/glossary-vba.md#target) position is relative to the left edge of the control. If the length of a line is less than the value of the **CurTargetX** property, you can place the insertion point at the end of the line. The value of **CurTargetX** changes when the user sets the insertion point or when the **CurX** property is set. **CurTargetX** is read-only.

The return value is valid when the object has [focus](../../Glossary/vbe-glossary.md#focus).

You can use **CurTargetX** and **CurX** to move the insertion point as the user scrolls through the contents of a multiline **TextBox** or **ComboBox**. When the user moves the insertion point to another line of text by scrolling the content of the object, **CurTargetX** specifies the preferred position for the insertion point. **CurX** is set to this value if the line of text is longer than the value of **CurTargetX**. Otherwise, **CurX** is set to the end of the line of text.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]