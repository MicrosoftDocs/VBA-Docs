---
title: Alignment property
keywords: fm20.chm5225002
f1_keywords:
- fm20.chm5225002
ms.prod: office
api_name:
- Office.Alignment
ms.assetid: d4c84882-dae6-ed8c-6dda-f042f22140cc
ms.date: 11/15/2018
localization_priority: Normal
---


# Alignment property

Specifies the position of a control relative to its caption.

## Syntax

_object_.**Alignment** [= _fmAlignment_ ]

The **Alignment** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmAlignment_|Optional. Caption position.|

## Settings

The settings for _fmAlignment_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmAlignmentLeft_|0|Places the caption to the left of the control.|
| _fmAlignmentRight_|1|Places the caption to the right of the control (default).|

## Remarks

The caption text for a control is left-aligned.

> [!NOTE] 
> Although the **Alignment** property exists on the **[ToggleButton](togglebutton-control.md)**, the property is disabled. You cannot set or return a value for this property on the **ToggleButton**.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]