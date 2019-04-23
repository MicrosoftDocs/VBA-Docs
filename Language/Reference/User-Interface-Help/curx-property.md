---
title: CurX property
keywords: fm20.chm2001040
f1_keywords:
- fm20.chm2001040
ms.prod: office
api_name:
- Office.CurX
ms.assetid: cbb6c8e9-13f2-61e7-9577-ceeef71ca2be
ms.date: 11/15/2018
localization_priority: Normal
---


# CurX property

Specifies the current horizontal position of the insertion point in a multiline **[TextBox](textbox-control.md)** or **[ComboBox](combobox-control.md)**.

## Syntax

_object_.**CurX** [= _Long_ ]

The **CurX** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. Indicates the current position, measured in himetrics. A himetric is 0.0001 meter.|

## Remarks

The **CurX** property applies to a multiline **TextBox** or **ComboBox**. The return value is valid when the object has the [focus](../../Glossary/vbe-glossary.md#focus).

You can use **CurTargetX** and **CurX** to position the insertion point as the user scrolls through the contents of a multiline **TextBox** or **ComboBox**. When the user moves the insertion point to another line of text by scrolling the content of the object, **CurTargetX** specifies the preferred position for the insertion point. 

**CurX** is set to this value if the line of text is longer than the value of **CurTargetX**. Otherwise, **CurX** is set to the end of the line of text.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]