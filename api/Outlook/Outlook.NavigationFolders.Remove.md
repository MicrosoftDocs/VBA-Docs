---
title: NavigationFolders.Remove Method (Outlook)
keywords: vbaol11.chm2898
f1_keywords:
- vbaol11.chm2898
ms.prod: outlook
api_name:
- Outlook.NavigationFolders.Remove
ms.assetid: ddaa3dd8-7539-ea5b-78a8-daa48ea63771
ms.date: 06/08/2017
---


# NavigationFolders.Remove Method (Outlook)

Removes an object from the collection.


## Syntax

 _expression_ . **Remove**( **_RemovableFolder_** )

 _expression_ A variable that represents a **[NavigationFolders](Outlook.NavigationFolders.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RemovableFolder_|Required| **[NavigationFolder](Outlook.NavigationFolder.md)**|The navigation folder to be removed.|

## Remarks

Only removable folders,  **NavigationFolder** objects with an **[IsRemovable](Outlook.NavigationFolder.IsRemovable.md)** property value set to **True** , can be removed from a **NavigationFolders** collection. This means that you can use **NavigationFolders.Remove** to remove shared folders, public folders, and linked folders. However, you must use **[Folder.Delete](Outlook.Folder.Delete.md)** to remove any user-created folders.


## See also


#### Concepts


[NavigationFolders Object](Outlook.NavigationFolders.md)

