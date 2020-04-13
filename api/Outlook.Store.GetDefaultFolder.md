---
title: Store.GetDefaultFolder method (Outlook)
keywords: vbaol11.chm3437
f1_keywords:
- vbaol11.chm3437
ms.prod: outlook
api_name:
- Outlook.Store.GetDefaultFolder
ms.assetid: f3e87528-6de8-dc59-8d27-f19f6b344044
ms.date: 06/08/2017
localization_priority: Normal
---


# Store.GetDefaultFolder method (Outlook)

Returns a **[Folder](Outlook.Folder.md)** object that represents the default folder in the store and that is of the type specified by the _FolderType_ argument.


## Syntax

_expression_. `GetDefaultFolder`( `_FolderType_` )

_expression_ A variable that represents a '[Store](Outlook.Store.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FolderType_|Required| **[OlDefaultFolders](Outlook.OlDefaultFolders.md)**|Specifies the type of the requested default folder.|

## Return value

A **Folder** object that represents the default folder of the requested type. If the default folder of the requested type does not exist, **GetDefaultFolder** returns **Null** (**Nothing** in Visual Basic).


## Remarks

This method is similar to the  **[GetDefaultFolder](Outlook.NameSpace.GetDefaultFolder.md)** method of the **[NameSpace](Outlook.NameSpace.md)** object. The difference is that this method gets the default folder on the delivery store that is associated with the account, whereas **NameSpace.GetDefaultFolder** returns the default folder on the default store for the current profile.

One example of when  **GetDefaultFolder** returns **Null** (**Nothing** in Visual Basic) is when **olFolderManagedEmail** is specified as the _FolderType_ but the Managed Folders group has not been deployed.


## See also


[Store Object](Outlook.Store.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]