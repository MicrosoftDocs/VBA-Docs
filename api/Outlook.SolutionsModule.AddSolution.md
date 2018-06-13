---
title: SolutionsModule.AddSolution Method (Outlook)
keywords: vbaol11.chm3368
f1_keywords:
- vbaol11.chm3368
ms.prod: outlook
api_name:
- Outlook.SolutionsModule.AddSolution
ms.assetid: 81d2edab-f8b3-340b-47b3-e98e780294ff
ms.date: 06/08/2017
---


# SolutionsModule.AddSolution Method (Outlook)

Adds a solution root folder and its subfolders to the  **Solutions** module.


## Syntax

 _expression_ . **AddSolution**( **_Solution_** , **_Scope_** )

 _expression_ A variable that represents a **[SolutionsModule](Outlook.SolutionsModule.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Solution_|Required| **[Folder](Outlook.Folder.md)**|Specifies the solution root folder to add to the  **Solutions** module.|
| _Scope_|Required| **[OlSolutionScope](Outlook.OlSolutionScope.md)**|Specifies whether to display the folders that are in the solution only in the  **Solutions** module and the **Folder List**, or to display them in their respective default modules in the Navigation Pane as well.|

## Remarks

If the  **AddSolution** method succeeds and no solution root folder previously existed under the **Solutions** module, Microsoft Outlook displays the **Solutions** module in the NavigationPane.

You cannot add the following folders to the  **Solutions** module as a solution root folder:


- A folder that Outlook displays on the Navigation Pane, as defined by the  **[OlDefaultFolders](Outlook.OlDefaultFolders.md)** enumeration.
    
- A special folder, as defined by the  **[OlSpecialFolders](Outlook.OlSpecialFolders.md)** enumeration.
    
- Any folder on a Microsoft Exchange Server public folder store. The  **[ExchangeStoreType](Outlook.Store.ExchangeStoreType.md)** property on the **[Store](Outlook.Folder.Store.md)** object for this folder is **olExchangePublicFolder** .
    
- A hidden folder. A hidden folder is one whose MAPI property,  **PR_ATTR_HIDDEN** , is **True** or one that is not in the IPM Subtree.
    


This method also returns an error if the folder that you specify already exists as a root folder or a subfolder in the  **Solutions** module, or if the folder is a parent folder of a folder in the **Solutions** module.

If the  _Scope_ parameter is set to **olShowInDefaultModules** of the **OlSolutionScope** enumeration, the solution root and its subfolders are displayed in their respective default modules as well as the **Solutions** module. If the _Scope_ parameter is set to **olHideInDefaultModules** , the solution root and its subfolders are displayed in the **Solutions** module.

Solution folders are always displayed in the  **Folder List** module.

By default, Outlook displays the  **Solutions** module after the **Tasks** module, provided that the navigation modules are in the default order ? **Mail**,  **Calendar**,  **Contacts**, and  **Tasks**. If the Navigation Pane is expanded, the  **Solutions** module is also initially displayed as an expanded module. If the **Tasks** module is not displayed, the **Solutions** module is displayed after the last expanded module in the Navigation Pane.


## See also


#### Concepts


[SolutionsModule Object](Outlook.SolutionsModule.md)

