---
title: TakeFocusOnClick property
keywords: fm20.chm5225047
f1_keywords:
- fm20.chm5225047
ms.prod: office
api_name:
- Office.TakeFocusOnClick
ms.assetid: 79768a90-398b-3224-0850-eb5a236eed7b
ms.date: 11/16/2018
localization_priority: Normal
---


# TakeFocusOnClick property

Specifies whether a control takes the [focus](../../Glossary/vbe-glossary.md#focus) when clicked.

## Syntax

_object_.**TakeFocusOnClick** [= _Boolean_ ]

The **TakeFocusOnClick** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Specifies whether a control takes the focus when clicked.|

## Settings

The settings for _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|The button takes the focus when clicked (default).|
|**False**|The button does not take the focus when clicked.|

## Remarks

The **TakeFocusOnClick** property defines only what happens when the user clicks a control. If the user tabs to the control, the control takes the focus regardless of the value of **TakeFocusOnClick**.

Use this property to complete actions that affect a control without requiring that control to give up focus. 

For example, assume that your form includes a **[TextBox](textbox-control.md)** and a **[CommandButton](commandbutton-control.md)** that checks for correct spelling of text. You would like to be able to select text in the **TextBox**, and then click the **CommandButton** and run the spelling checker without taking focus away from the **TextBox**. You can do this by setting the **TakeFocusOnClick** property of the **CommandButton** to **False**.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]