---
title: MultiSelect property (Microsoft Forms)
keywords: fm20.chm5225069
f1_keywords:
- fm20.chm5225069
ms.prod: office
ms.assetid: 4c8102d4-abbb-a7f7-8dd3-0a0695752fa8
ms.date: 11/16/2018
localization_priority: Normal
---


# MultiSelect property (Microsoft Forms)

Indicates whether the object permits multiple selections.

## Syntax

_object_.**MultiSelect** [= _fmMultiSelect_ ]

The **MultiSelect** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmMultiSelect_|Optional. The selection mode that the control uses.|

## Settings

The settings for _fmMultiSelect_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmMultiSelectSingle_|0|Only one item can be selected (default).|
| _fmMultiSelectMulti_|1|Pressing the SPACEBAR or clicking selects or deselects an item in the list.|
| _fmMultiSelectExtended_|2|Pressing SHIFT and clicking the mouse, or pressing SHIFT and one of the arrow keys, extends the selection from the previously selected item to the current item.<br/><br/>Pressing CTRL and clicking the mouse selects or deselects an item.|

## Remarks

When the **MultiSelect** property is set to _Extended_ or _Multi_, you must use the list box's **Selected** property to determine the selected items. Also, the **Value** property of the control is always **Null**.

The **ListIndex** property returns the index of the row with the keyboard [focus](../../Glossary/vbe-glossary.md#focus).

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]