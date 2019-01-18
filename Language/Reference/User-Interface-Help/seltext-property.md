---
title: SelText property
keywords: fm20.chm2001890
f1_keywords:
- fm20.chm2001890
ms.prod: office
api_name:
- Office.SelText
ms.assetid: 75b9c27f-f6f7-6445-6d86-a53f046c1db6
ms.date: 11/16/2018
localization_priority: Normal
---


# SelText property

Returns or sets the selected text of a control.

## Syntax

_object_.**SelText** [= _String_ ]

The **SelText** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. A string expression containing the selected text.|

## Remarks

If no characters are selected in the edit region of the control, the **SelText** property returns a zero length string. This property is valid regardless of whether the control has the [focus](../../Glossary/vbe-glossary.md#focus).

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]