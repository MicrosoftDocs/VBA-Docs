---
title: WordWrap property
keywords: fm20.chm5225114
f1_keywords:
- fm20.chm5225114
ms.prod: office
api_name:
- Office.WordWrap
ms.assetid: c68f3da4-d930-62cc-b9fb-5f2de42d413f
ms.date: 11/16/2018
localization_priority: Normal
---


# WordWrap property

Indicates whether the contents of a control automatically wrap at the end of a line.

## Syntax

_object_.**WordWrap** [= _Boolean_ ]

The **WordWrap** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the control expands to fit the text.|

## Settings

The settings for  _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|The text wraps (default).|
|**False**|The text does not wrap.|

## Remarks

For controls that support the **MultiLine** property as well as the **WordWrap** property, **WordWrap** is ignored when **MultiLine** is **False**.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]