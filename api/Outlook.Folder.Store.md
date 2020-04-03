---
title: Folder.Store property (Outlook)
keywords: vbaol11.chm2016
f1_keywords:
- vbaol11.chm2016
ms.prod: outlook
api_name:
- Outlook.Folder.Store
ms.assetid: 347d3031-01cf-a248-4abc-f749feb811a4
ms.date: 06/08/2017
localization_priority: Normal
---


# Folder.Store property (Outlook)

Returns a **[Store](Outlook.Store.md)** object representing the store that contains the **[Folder](Outlook.Folder.md)** object. Read-only.


## Syntax

_expression_. `Store`

_expression_ A variable that represents a [Folder](Outlook.Folder.md) object.


## Remarks

This property returns a **Store** object except in the case where the **Folder** is a shared folder (returned by **[NameSpace.GetSharedDefaultFolder](Outlook.NameSpace.GetSharedDefaultFolder.md)**). In this case, one user has delegated access to a default folder to another user; a call to **Folder.Store** will return **Null**.


## See also


[Folder Object](Outlook.Folder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]