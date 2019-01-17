---
title: File Attribute constants
keywords: vblr6.chm1113237
f1_keywords:
- vblr6.chm1113237
ms.prod: office
ms.assetid: ed8e634c-bb35-f1e4-af35-593ced56d997
ms.date: 12/11/2018
localization_priority: Normal
---


# File Attribute constants

These constants are only available when your project has an explicit reference to the appropriate [type library](../../Glossary/vbe-glossary.md#type-library) containing these constant definitions.

<br/>

|Constant|Value|Description|
|:-----|:-----|:-----|
|**Normal**|0|Normal file. No attributes are set.|
|**ReadOnly**|1|Read-only file. Attribute is read/write.|
|**Hidden**|2|Hidden file. Attribute is read/write.|
|**System**|4|System file. Attribute is read/write.|
|**Volume**|8|Disk drive volume label. Attribute is read-only.|
|**Directory**|16|Folder or directory. Attribute is read-only.|
|**Archive**|32|File has changed since last backup. Attribute is read/write.|
|**Alias**|64|Link or shortcut. Attribute is read-only.|
|**Compressed**|128|Compressed file. Attribute is read-only.|

## See also

- [Constants (Visual Basic for Applications)](../constants-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]