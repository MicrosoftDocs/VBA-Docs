---
title: AutoTab property
keywords: fm20.chm2000750
f1_keywords:
- fm20.chm2000750
ms.prod: office
api_name:
- Office.AutoTab
ms.assetid: 36af6755-72d8-439a-2999-fc4224760529
ms.date: 11/15/2018
localization_priority: Normal
---


# AutoTab property

Specifies whether an automatic tab occurs when a user enters the maximum allowable number of characters into a **[TextBox](textbox-control.md)** or the text box portion of a **[ComboBox](combobox-control.md)**.

## Syntax

_object_.**AutoTab** [= _Boolean_ ]

The **AutoTab** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Specifies whether an automatic tab occurs.|

## Settings

The settings for _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|Tab occurs.|
|**False**|Tab does not occur (default).|

## Remarks

The **MaxLength** property specifies the maximum number of characters allowed in a **TextBox** or the text box portion of a **ComboBox**.

You can specify the **AutoTab** property for a **TextBox** or **ComboBox** on a form for which you usually enter a set number of characters. Once a user enters the maximum number of characters, the [focus](../../Glossary/vbe-glossary.md#focus) automatically moves to the next control in the [tab order](../../Glossary/vbe-glossary.md#tab-order). For example, if a **TextBox** displays inventory stock numbers that are always five characters long, you can use **MaxLength** to specify the maximum number of characters to enter into the **TextBox** and **AutoTab** to automatically tab to the next control after the user enters five characters.

Support for **AutoTab** varies from one application to another. Not all [containers](../../Glossary/vbe-glossary.md#container) support this property.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]