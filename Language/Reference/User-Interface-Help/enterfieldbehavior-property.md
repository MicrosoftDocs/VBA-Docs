---
title: EnterFieldBehavior property
keywords: fm20.chm5225036
f1_keywords:
- fm20.chm5225036
ms.prod: office
api_name:
- Office.EnterFieldBehavior
ms.assetid: 6657b5c5-d204-1c5e-c8d7-e84bc51efe15
ms.date: 11/16/2018
localization_priority: Normal
---


# EnterFieldBehavior property

Specifies the selection behavior when entering a **[TextBox](textbox-control.md)** or **[ComboBox](combobox-control.md)**.

## Syntax

_object_.**EnterFieldBehavior** [= _fmEnterFieldBehavior_ ]

The **EnterFieldBehavior** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmEnterFieldBehavior_|Optional. The desired selection behavior.|

## Settings

The settings for _fmEnterFieldBehavior_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmEnterFieldBehaviorSelectAll_|0|Selects the entire contents of the edit region when entering the control (default).|
| _fmEnterFieldBehaviorRecallSelection_|1|Leaves the selection unchanged. Visually, this uses the selection that was in effect the last time the control was active.|

## Remarks

The **EnterFieldBehavior** property controls the way text is selected when the user tabs to the control, not when the control receives [focus](../../Glossary/vbe-glossary.md#focus) as a result of the **SetFocus** method. 

Following **SetFocus**, the contents of the control are not selected, and the insertion point appears after the last character in the control's edit region.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]