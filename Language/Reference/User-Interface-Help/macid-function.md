---
title: MacID function (Visual Basic for Applications)
keywords: vblr6.chm1009753
f1_keywords:
- vblr6.chm1009753
ms.prod: office
ms.assetid: 4df07ec9-c165-ab5a-2864-ef1d9168d7e5
ms.date: 12/13/2018
localization_priority: Normal
---


# MacID function

Used on the Macintosh to convert a 4-character [constant](../../Glossary/vbe-glossary.md#constant) to a value that may be used by **Dir**, **Kill**, **Shell**, and **AppActivate**.

## Syntax

**MacID**(_constant_)

The required _constant_ argument consists of 4 characters used to specify a resource type, file type, application signature, or Apple Event; for example, TEXT, OBIN, "XLS5" for Excel files ("XLS8" for Excel 97); Microsoft Word uses "W6BN" ("W8BN" for Word 97), and so on.

## Remarks

**MacID** is used with **Dir** and **Kill** to specify a Macintosh file type. Because the Macintosh does not support **\*** and **?** as wildcards, you can use a four-character constant instead to identify groups of files. For example, the following statement returns TEXT type files from the current folder:

```vb
Dir("SomePath", MacID("TEXT"))

```

**MacID** is used with **Shell** and **AppActivate** to specify an application by using the application's unique signature.

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]