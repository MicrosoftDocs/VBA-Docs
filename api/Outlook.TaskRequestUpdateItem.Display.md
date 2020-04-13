---
title: TaskRequestUpdateItem.Display method (Outlook)
keywords: vbaol11.chm1950
f1_keywords:
- vbaol11.chm1950
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.Display
ms.assetid: d91b4860-12f5-b898-9717-b4fddfaee49f
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.Display method (Outlook)

Displays a new **[Inspector](Outlook.Inspector.md)** object for the item.


## Syntax

_expression_. `Display`( `_Modal_` )

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Modal_|Optional| **Variant**| **True** to make the window modal. The default value is **False**.|

## Remarks

The **Display** method is supported for explorer and inspector windows for the sake of backward compatibility. To activate an explorer or inspector window, use the **[Activate](Outlook.Inspector.Activate(method).md)** method.

If you attempt to open an "unsafe" file system object (or "freedoc" file) by using the Microsoft Outlook object model, you receive the  **E_FAIL** return code in the C or C++ programming languages. In Outlook 2000 and earlier, you could open an "unsafe" file system object by using the **Display** method.


## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]