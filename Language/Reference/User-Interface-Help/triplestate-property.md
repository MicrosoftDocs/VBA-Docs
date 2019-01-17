---
title: TripleState property
keywords: fm20.chm5225110
f1_keywords:
- fm20.chm5225110
ms.prod: office
api_name:
- Office.TripleState
ms.assetid: f009f524-76db-526f-7bd6-a7358b53fc31
ms.date: 11/16/2018
localization_priority: Normal
---


# TripleState property

Determines whether a user can specify, from the user interface, the [Null](../../Glossary/vbe-glossary.md#null) state for a **[CheckBox](checkbox-control.md)** or **[ToggleButton](togglebutton-control.md)**.

## Syntax

_object_.**TripleState** [= _Boolean_ ]

The **TripleState** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the control supports the Null state.|

## Settings

The settings for _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|The button clicks through three states.|
|**False**|The button only supports True and False (default).|

## Remarks

Although the **TripleState** property exists on the **[OptionButton](optionbutton-control.md)**, the property is disabled. Regardless of the value of **TripleState**, you cannot set the control to **Null** through the user interface.

When the **TripleState** property is **True**, a user can choose from the values of **Null**, **True**, and **False**. The null value is displayed as a shaded button.

When **TripleState** is **False**, the user can choose either **True** or **False**.

A control set to **Null** does not initiate the Click event.

Regardless of the property setting, the **Null** value can always be assigned programmatically to a **CheckBox** or **ToggleButton**, causing that control to appear shaded.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]