---
title: SelLength property
keywords: fm20.chm2001870
f1_keywords:
- fm20.chm2001870
ms.prod: office
api_name:
- Office.SelLength
ms.assetid: 86f86e84-b22e-a86a-12b9-dc1011cbcf9d
ms.date: 11/16/2018
localization_priority: Normal
---


# SelLength property

The number of characters selected in a **[TextBox](textbox-control.md)** or the text portion of a **[ComboBox](combobox-control.md)**.

## Syntax

_object_.**SelLength** [= _Long_ ]

The **SelLength** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. A numeric expression specifying the number of characters selected.<br/><br/>For **SelLength** and **SelStart**, the valid range of settings is 0 to the total number of characters in the edit area of a **ComboBox** or **TextBox**.|

## Remarks

The **SelLength** property is always valid, even when the control does not have [focus](../../Glossary/vbe-glossary.md#focus). 

Setting **SelLength** to a value less than zero creates an error. Attempting to set **SelLength** to a value greater than the number of characters available in a control results in a value equal to the number of characters in the control.

> [!NOTE] 
> Changing the value of the **SelStart** property cancels any existing selection in the control, places an insertion point in the text, and sets **SelLength** to zero.

The default value, zero, means that no text is currently selected.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]