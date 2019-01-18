---
title: Orientation property (Microsoft Forms)
keywords: fm20.chm5225074
f1_keywords:
- fm20.chm5225074
ms.prod: office
ms.assetid: 3e57f9af-8aa5-85f5-f3af-81f9a61373c0
ms.date: 11/16/2018
localization_priority: Normal
---


# Orientation property (Microsoft Forms)

Specifies whether the **[SpinButton](spinbutton-control.md)** or **[ScrollBar](scrollbar-control.md)** is oriented vertically or horizontally.

## Syntax

_object_.**Orientation** [= _fmOrientation_ ]

The **Orientation** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmOrientation_|Optional. Orientation of the control.|

## Settings

The settings for _fmOrientation_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmOrientationAuto_|-1|Automatically determines the orientation based upon the dimensions of the control (default).|
| _FmOrientationVertical_|0|Control is rendered vertically.|
| _FmOrientationHorizontal_|1|Control is rendered horizontally.|

## Remarks

If you specify automatic orientation, the height and width of the control determine whether it appears horizontally or vertically. For example, if the control is wider than it is tall, it appears horizontally; if it is taller than it is wide, the control appears vertically.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]