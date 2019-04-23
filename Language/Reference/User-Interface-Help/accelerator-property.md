---
title: Accelerator property
keywords: fm20.chm2000690
f1_keywords:
- fm20.chm2000690
ms.prod: office
api_name:
- Office.Accelerator
ms.assetid: d9183848-4638-745b-e3f4-b076493d3668
ms.date: 11/15/2018
localization_priority: Normal
---


# Accelerator property

Sets or retrieves the [accelerator key](../../Glossary/glossary-vba.md#accelerator-key) for a control.

## Syntax

_object_.**Accelerator** [= _String_ ]

The **Accelerator** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. The character to use as the accelerator key.|

## Remarks

To designate an accelerator key, enter a single character for the **Accelerator** property. You can set **Accelerator** in the control's property sheet or in code. If the value of this property contains more than one character, the first character in the string becomes the value of **Accelerator**.

When an accelerator key is used, there is no visual feedback (other than [focus](../../Glossary/vbe-glossary.md#focus)) to indicate that the control initiated the Click event. For example, if the accelerator key applies to a **[CommandButton](commandbutton-control.md)**, the user will not see the button pressed in the interface. The button receives the focus, however, when the user presses the accelerator key.

If the accelerator applies to a **[Label](label-control.md)**, the control following the **Label** in the [tab order](../../Glossary/vbe-glossary.md#tab-order), rather than the **Label** itself, receives the focus.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]