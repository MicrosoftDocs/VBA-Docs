---
title: AutoWordSelect property
keywords: fm20.chm2000760
f1_keywords:
- fm20.chm2000760
ms.prod: office
api_name:
- Office.AutoWordSelect
ms.assetid: 24e9e8ff-5988-9ed3-4a2c-f3faa99248f9
ms.date: 11/15/2018
localization_priority: Normal
---


# AutoWordSelect property

Specifies whether a word or a character is the basic unit used to extend a selection.

## Syntax

_object_.**AutoWordSelect** [= _Boolean_ ]

The **AutoWordSelect** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Specifies the basic unit used to extend a selection.|

## Settings

The settings for _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|Uses a word as the basic unit (default).|
|**False**|Uses a character as the basic unit.|

## Remarks

The **AutoWordSelect** property specifies how the selection extends or contracts in the edit region of a **[TextBox](textbox-control.md)** or **[ComboBox](combobox-control.md)**.

If the user places the insertion point in the middle of a word and then extends the selection while **AutoWordSelect** is **True**, the selection includes the entire word.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]