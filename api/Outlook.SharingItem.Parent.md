---
title: SharingItem.Parent property (Outlook)
keywords: vbaol11.chm596
f1_keywords:
- vbaol11.chm596
ms.prod: outlook
api_name:
- Outlook.SharingItem.Parent
ms.assetid: 78d6d287-9623-0ed0-eab6-75a0a57d0c6c
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.Parent property (Outlook)

Returns the parent  **Object** of the specified **[SharingItem](Outlook.SharingItem.md)**. Read-only.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

If the  **SharingItem** was just created, this property returns a **[Folder](Outlook.Folder.md)** object representing the **Inbox** folder. Otherwise, this property returns a **Folder** object representing the folder in which the **SharingItem** was saved.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]