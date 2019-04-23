---
title: VerticalScrollbarSide property
keywords: fm20.chm5225112
f1_keywords:
- fm20.chm5225112
ms.prod: office
api_name:
- Office.VerticalScrollbarSide
ms.assetid: 0439743b-3774-5778-7022-dbeea5ef8c39
ms.date: 11/16/2018
localization_priority: Normal
---


# VerticalScrollbarSide property

Specifies whether a vertical scroll bar appears on the right or left side of a form or page.

## Syntax

_object_.**VerticalScrollbarSide** [= _fmVerticalScrollbarSide_ ]

The **VerticalScrollbarSide** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmVerticalScrollbarSide_|Optional. Where the scroll bar should appear.|

## Settings

The settings for _fmVerticalScrollbarSide_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmVerticalScrollbarSideRight_|0|Puts the scroll bar on the right side (default).|
| _fmVerticalScrollBarSideLeft_|1|Puts the scroll bar on the left side.|

## Remarks

The **VerticalScrollBarSide** property is particularly useful if the form will be used in an environment where reading occurs from right to left.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]