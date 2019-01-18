---
title: WhatsThisButton property (Visual Basic for Applications)
keywords: vblr6.chm916694
f1_keywords:
- vblr6.chm916694
ms.prod: office
api_name:
- Office.WhatsThisButton
ms.assetid: f9e24796-d4e0-1719-32b3-2119f20a6b5a
ms.date: 12/19/2018
localization_priority: Normal
---


# WhatsThisButton property

Returns a [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value that determines whether the **What's This** button appears on the title bar of a **[UserForm](userform-object.md)** object. Read-only at [run time](../../Glossary/vbe-glossary.md#run-time). This property does not apply to the Macintosh.

## Remarks

The settings for the **WhatsThisButton** property are:

|Setting|Description|
|:-----|:-----|
|**True**|Turns on display of the **What's This Help** button.|
|**False**|(Default) Turns off display of the **What's This Help** button.|


## Remarks

The **WhatsThisHelp** property must be **True** for the **WhatsThisButton** property to be **True**.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]