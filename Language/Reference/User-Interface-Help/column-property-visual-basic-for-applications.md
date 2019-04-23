---
title: Column property (Visual Basic for Applications)
keywords: vblr6.chm2182073
f1_keywords:
- vblr6.chm2182073
ms.prod: office
ms.assetid: 5733f4a5-cf81-632f-8a29-df71951d0c7e
ms.date: 12/19/2018
localization_priority: Normal
---


# Column property

Read-only property that returns the column number of the current character position in a **TextStream** file.

## Syntax

_object_.**Column**

The _object_ is always the name of a **[TextStream](textstream-object.md)** object.

## Remarks

After a newline character has been written, but before any other character is written, **Column** is equal to 1.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]