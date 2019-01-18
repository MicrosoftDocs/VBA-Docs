---
title: ShowDropButtonWhen property
keywords: fm20.chm2001900
f1_keywords:
- fm20.chm2001900
ms.prod: office
api_name:
- Office.ShowDropButtonWhen
ms.assetid: 82c7a038-a4fa-7253-ec24-c97e6841293e
ms.date: 11/16/2018
localization_priority: Normal
---


# ShowDropButtonWhen property

Specifies when to show the drop-down button for a **[ComboBox](combobox-control.md)** or **[TextBox](textbox-control.md)**.

## Syntax

_object_.**ShowDropButtonWhen** [= _fmShowDropButtonWhen_ ]

The **ShowDropButtonWhen** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmShowDropButtonWhen_|Optional. The circumstances under which the drop-down button will be visible.|

## Settings

The settings for _fmShowDropButtonWhen_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmShowDropButtonWhenNever_|0|Do not show the drop-down button under any circumstances.|
| _fmShowDropButtonWhenFocus_|1|Show the drop-down button when the control has the focus.|
| _fmShowDropButtonWhenAlways_|2|Always show the drop-down button.|

## Remarks

For a **ComboBox**, the default value is _fmShowDropButtonWhenAlways_.

For a **TextBox**, the default value is _fmShowDropButtonWhenNever_.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]