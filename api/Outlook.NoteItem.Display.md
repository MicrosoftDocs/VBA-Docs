---
title: NoteItem.Display method (Outlook)
keywords: vbaol11.chm1495
f1_keywords:
- vbaol11.chm1495
ms.prod: outlook
api_name:
- Outlook.NoteItem.Display
ms.assetid: 1a8c3999-45d4-b4a1-dacf-371a7e711eb2
ms.date: 06/08/2017
localization_priority: Normal
---


# NoteItem.Display method (Outlook)

Displays a new **[Inspector](Outlook.Inspector.md)** object for the item.


## Syntax

_expression_. `Display`( `_Modal_` )

_expression_ A variable that represents a [NoteItem](Outlook.NoteItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Modal_|Optional| **Variant**| **True** to make the window modal. The default value is **False**.|

## Remarks

The **Display** method is supported for explorer and inspector windows for the sake of backward compatibility. To activate an explorer or inspector window, use the **[Activate](Outlook.Inspector.Activate(method).md)** method.

If you attempt to open an "unsafe" file system object (or "freedoc" file) by using the Microsoft Outlook object model, you receive the  **E_FAIL** return code in the C or C++ programming languages. In Outlook 2000 and earlier, you could open an "unsafe" file system object by using the **Display** method.


## See also


[NoteItem Object](Outlook.NoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]