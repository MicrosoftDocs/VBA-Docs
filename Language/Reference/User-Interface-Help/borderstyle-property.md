---
title: BorderStyle property
keywords: fm20.chm5225010
f1_keywords:
- fm20.chm5225010
ms.prod: office
api_name:
- Office.BorderStyle
ms.assetid: 211ffd49-cf3a-8fff-1f00-58a97b833580
ms.date: 11/15/2018
localization_priority: Normal
---


# BorderStyle property

Specifies the type of border used by a control or a form.

## Syntax

_object_.**BorderStyle** [= _fmBorderStyle_ ]

The **BorderStyle** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmBorderStyle_|Optional. Specifies the border style.|

## Settings

The settings for _fmBorderStyle_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmBorderStyleNone_|0|The control has no visible border line.|
| _fmBorderStyleSingle_|1|The control has a single-line border (default).|

The default value for a **[ComboBox](combobox-control.md)**, **[Frame](frame-control.md)**, **[Label](label-control.md)**, **[ListBox](listbox-control.md)**, or **[TextBox](textbox-control.md)** is 0 ( _None_ ). The default value for an **[Image](image-control.md)** is 1 ( _Single_ ).

## Remarks

For a **Frame**, the **BorderStyle** property is ignored if the **SpecialEffect** property is _None_.

You can use either **BorderStyle** or **SpecialEffect** to specify the border for a control, but not both. If you specify a nonzero value for one of these properties, the system sets the value of the other property to zero. For example, if you set **BorderStyle** to **fmBorderStyleSingle**, the system sets **SpecialEffect** to zero ( _Flat_ ). If you specify a nonzero value for **SpecialEffect**, the system sets **BorderStyle** to zero.

**BorderStyle** uses **BorderColor** to define the colors of its borders.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]