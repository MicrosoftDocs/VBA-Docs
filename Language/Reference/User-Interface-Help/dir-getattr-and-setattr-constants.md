---
title: Dir, GetAttr, and SetAttr constants
keywords: vblr6.chm1012529
f1_keywords:
- vblr6.chm1012529
ms.prod: office
ms.assetid: ca85f083-4824-1371-238b-f1ac55f8f702
ms.date: 12/11/2018
localization_priority: Normal
---


# Dir, GetAttr, and SetAttr constants

The following [constants](../../Glossary/vbe-glossary.md#constant) can be used anywhere in your code in place of the actual values.

<br/>

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbNormal**|0|Normal (default for **Dir** and **SetAttr**)|
|**vbReadOnly**|1|Read-only|
|**vbHidden**|2|Hidden|
|**vbSystem**|4|System file|
|**vbVolume**|8|Volume label|
|**vbDirectory**|16|Directory or folder|
|**vbArchive**|32|File has changed since last backup|
|**vbAlias**|64|On the Macintosh, identifier is an alias|

Only **VbNormal**, **vbReadOnly**, **vbHidden**, and **vbAlias** are available on the Macintosh.

## See also

- [Dir function](dir-function.md)
- [GetAttr function](getattr-function.md)
- [SetAttr statement](setattr-statement.md)
- [Constants (Visual Basic for Applications)](../constants-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]