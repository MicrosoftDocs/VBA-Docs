---
title: Description property (Visual Basic for Applications)
keywords: vblr6.chm1014191
f1_keywords:
- vblr6.chm1014191
ms.prod: office
ms.assetid: cab35a69-b45a-2d96-f495-2fae208fca6a
ms.date: 12/19/2018
localization_priority: Normal
---


# Description property

Returns or sets a [string expression](../../Glossary/vbe-glossary.md#string-expression) containing a descriptive string associated with an object. Read/write.

For the **[Err](err-object.md)** object, returns or sets a descriptive string associated with an error.

## Remarks

The **Description** property setting consists of a short description of the error. Use this [property](../../Glossary/vbe-glossary.md#property) to alert the user to an error that you either can't or don't want to handle. 

When generating a user-defined error, assign a short description of your error to the **Description** property. If **Description** isn't filled in, and the value of **[Number](number-property-visual-basic-for-applications.md)** corresponds to a Visual Basic [run-time error](../../Glossary/vbe-glossary.md#run-time-error), the string returned by the **[Error](error-function.md)** function is placed in **Description** when the error is generated.

## Example

This example assigns a user-defined message to the **Description** property of the **Err** object.

```vb
Err.Description = "It was not possible to access an object necessary " _
& "for this operation."

```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]