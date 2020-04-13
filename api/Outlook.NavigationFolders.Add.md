---
title: NavigationFolders.Add method (Outlook)
keywords: vbaol11.chm2897
f1_keywords:
- vbaol11.chm2897
ms.prod: outlook
api_name:
- Outlook.NavigationFolders.Add
ms.assetid: f88fd69a-8684-bfc4-bc20-1cff5c44974e
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationFolders.Add method (Outlook)

Adds the specified  **[Folder](Outlook.Folder.md)**, as a **[NavigationFolder](Outlook.NavigationFolder.md)** object, to the end of the **[NavigationFolders](Outlook.NavigationFolders.md)** collection.


## Syntax

_expression_.**Add** (_Folder_)

_expression_ A variable that represents a [NavigationFolders](Outlook.NavigationFolders.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Folder_|Required| **Folder**|The folder to add.|

## Return value

A **NavigationFolder** object that represents the new navigation folder.


## Remarks

A folder can only appear in one navigation group. When adding a **Folder** object to a new navigation group, any references to that **Folder** are removed from any other navigation group of which it was previously a member.


## See also


[NavigationFolders Object](Outlook.NavigationFolders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]