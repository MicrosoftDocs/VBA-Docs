---
title: SharingItem.AllowWriteAccess property (Outlook)
keywords: vbaol11.chm700
f1_keywords:
- vbaol11.chm700
ms.prod: outlook
api_name:
- Outlook.SharingItem.AllowWriteAccess
ms.assetid: 538c9681-d164-52ff-eb8b-4ae0c6875247
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.AllowWriteAccess property (Outlook)

Returns or sets a **Boolean** value that indicates whether a sharing invitation should include write access to the folder. Read/write.


## Syntax

_expression_. `AllowWriteAccess`

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Return value

 **True** if the recipient of the sharing invitation should receive write access; otherwise, **False**. The default is **False**.


## Remarks

When sending a sharing invitation for a non-default folder, the recipient can be granted write access to the folder in addition to the default read access. This property determines if write permission should be granted to the recipient when the  **[SharingItem](Outlook.SharingItem.md)** is sent.

An error occurs if you attempt to set this property after the sharing message has been sent or received.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]