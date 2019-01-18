---
title: SkipLine method (Visual Basic for Applications)
keywords: vblr6.chm2182080
f1_keywords:
- vblr6.chm2182080
ms.prod: office
api_name:
- Office.SkipLine
ms.assetid: 77015ee6-d778-d38b-5f5b-b18f65e828fd
ms.date: 12/14/2018
localization_priority: Normal
---


# SkipLine method

Skips the next line when reading a **TextStream** file.

## Syntax

_object_.**SkipLine**

The _object_ is always the name of a **[TextStream](textstream-object.md)** object.

## Remarks

Skipping a line means reading and discarding all characters in a line up to and including the next newline character.

An error occurs if the file is not open for reading.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]