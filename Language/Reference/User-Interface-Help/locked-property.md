---
title: Locked property
keywords: fm20.chm5225059
f1_keywords:
- fm20.chm5225059
ms.prod: office
api_name:
- Office.Locked
ms.assetid: 08bf09c4-0445-0749-daf2-a0fab8787ea8
ms.date: 11/16/2018
localization_priority: Normal
---


# Locked property

Specifies whether a control can be edited.

## Syntax

_object_.**Locked** [= _Boolean_ ]

The **Locked** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the control can be edited.|

## Settings

The settings for  _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|You can't edit the value.|
|**False**|You can edit the value (default).|

## Remarks

When a control is locked and enabled, it can still initiate events and can still receive the [focus](../../Glossary/vbe-glossary.md#focus).

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]