---
title: PasswordChar property
keywords: fm20.chm5225076
f1_keywords:
- fm20.chm5225076
ms.prod: office
api_name:
- Office.PasswordChar
ms.assetid: 2dd645b2-fe8d-a644-b796-e0595627cbb8
ms.date: 12/29/2018
localization_priority: Normal
---


# PasswordChar property

Specifies whether [placeholder](../../Glossary/glossary-vba.md#placeholder) characters are displayed instead of the characters actually entered in a **[TextBox](../../reference/user-interface-help/textbox-control.md)**.

## Syntax

_object_.**PasswordChar** [= _String_ ]

The **PasswordChar** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. A string expression specifying the placeholder character.|

## Remarks

You can use the **PasswordChar** property to protect sensitive information, such as passwords or security codes. The value of **PasswordChar** is the character that appears in a control instead of the actual characters that the user types. If you don't specify a character, the control displays the characters that the user types.

## See also

- [Microsoft Forms examples](../../reference/user-interface-help/examples-microsoft-forms.md)
- [Microsoft Forms reference](../../reference/user-interface-help/reference-microsoft-forms.md)
- [Microsoft Forms concepts](../../reference/user-interface-help/concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]