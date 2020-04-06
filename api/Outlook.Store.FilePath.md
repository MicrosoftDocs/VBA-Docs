---
title: Store.FilePath property (Outlook)
keywords: vbaol11.chm803
f1_keywords:
- vbaol11.chm803
ms.prod: outlook
api_name:
- Outlook.Store.FilePath
ms.assetid: 3b0ed312-9304-61a6-7152-5693a0e2f0fe
ms.date: 06/08/2017
localization_priority: Normal
---


# Store.FilePath property (Outlook)

Returns a  **String** representing the full file path for a Personal Folders File (.pst) or an Offline Folder File (.ost) store. Read-only.


## Syntax

_expression_. `FilePath`

_expression_ A variable that represents a [Store](Outlook.Store.md) object.


## Remarks

This property supports both Exchange and non-Exchange stores. If the store is not a .pst or .ost store,  **FilePath** returns an empty string.


## See also


[Store Object](Outlook.Store.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]