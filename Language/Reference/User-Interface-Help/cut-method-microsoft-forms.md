---
title: Cut method (Microsoft Forms)
keywords: fm20.chm2000290
f1_keywords:
- fm20.chm2000290
ms.prod: office
ms.assetid: 9eea6f19-557d-2ae0-4e22-2f40b4d01caf
ms.date: 11/15/2018
localization_priority: Normal
---


# Cut method (Microsoft Forms)

Removes selected information from an object and transfers it to the Clipboard.

## Syntax

_object_. **Cut**

The **Cut** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Remarks

For a **[ComboBox](combobox-control.md)** or **[TextBox](textbox-control.md)**, the **Cut** method removes currently selected text in the control to the Clipboard. This method does not require that the control have the [focus](../../Glossary/vbe-glossary.md#focus).

On a **Page**, **[Frame](frame-control.md)**, or form, **Cut** removes currently selected controls to the Clipboard. This method only removes controls created at [run time](../../Glossary/vbe-glossary.md#run-time).

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]