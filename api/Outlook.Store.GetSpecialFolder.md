---
title: Store.GetSpecialFolder method (Outlook)
keywords: vbaol11.chm812
f1_keywords:
- vbaol11.chm812
ms.prod: outlook
api_name:
- Outlook.Store.GetSpecialFolder
ms.assetid: 8f768a43-1589-5659-76f3-43afa4b745b6
ms.date: 06/08/2017
localization_priority: Normal
---


# Store.GetSpecialFolder method (Outlook)

Returns a **[Folder](Outlook.Folder.md)** object for a special folder specified by _FolderType_ in a given store.


## Syntax

_expression_. `GetSpecialFolder`( `_FolderType_` )

_expression_ A variable that represents a [Store](Outlook.Store.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FolderType_|Required| **[OlSpecialFolders](Outlook.OlSpecialFolders.md)**|A constant in the  **OlSpecialFolders** enumeration that specifies the type of the special folder in the store.|

## Return value

A **Folder** object that represents a special folder specified by the _FolderType_.


## Remarks

Not all special folders exist in all stores. If the requested special folder does not exist, a **Null** value (**Nothing** in VB) will be returned.


## See also


[Store Object](Outlook.Store.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]