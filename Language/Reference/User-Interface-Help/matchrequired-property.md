---
title: MatchRequired property
keywords: fm20.chm2001500
f1_keywords:
- fm20.chm2001500
ms.prod: office
api_name:
- Office.MatchRequired
ms.assetid: c2b2d308-4107-975f-9a2d-e0eaff413807
ms.date: 11/16/2018
localization_priority: Normal
---


# MatchRequired property

Specifies whether a value entered in the text portion of a **[ComboBox](combobox-control.md)** must match an entry in the existing list portion of the control. The user can enter non-matching values, but may not leave the control until a matching value is entered.

## Syntax

_object_.**MatchRequired** [= _Boolean_ ]

The **MatchRequired** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the text entered must match an existing item in the list.|

## Settings

The settings for _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|The text entered must match an existing list entry.|
|**False**|The text entered can be different from all existing list entries (default).|

## Remarks

If the **MatchRequired** property is **True**, the user cannot exit the **ComboBox** until the text entered matches an entry in the existing list. **MatchRequired** maintains the integrity of the list by requiring the user to select an existing entry.

> [!NOTE] 
> Not all [containers](../../Glossary/vbe-glossary.md#container) enforce this property.


## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]