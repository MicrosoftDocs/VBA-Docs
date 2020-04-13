---
title: NavigationGroups.Delete method (Outlook)
keywords: vbaol11.chm2859
f1_keywords:
- vbaol11.chm2859
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.Delete
ms.assetid: b5bb08c4-9cf1-4ed7-9522-0096f1016e5b
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationGroups.Delete method (Outlook)

Deletes the specified  **[NavigationGroup](Outlook.NavigationGroup.md)** object from the **[NavigationGroups](Outlook.NavigationGroups.md)** collection.


## Syntax

_expression_.**Delete**( `_Group_` )

_expression_ A variable that represents a [NavigationGroups](Outlook.NavigationGroups.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Group_|Required| **NavigationGroup**|The navigation group to be deleted.|

## Remarks

The **Delete** method raises an error if:


-  The navigation group specified in _Group_ contains navigation folders in its **[NavigationFolders](Outlook.NavigationFolders.md)** collection.
    
- The **[GroupType](Outlook.NavigationGroup.GroupType.md)** property of the navigation group specified in _Group_ is set to **olMyFoldersGroup**.
    
- The parent of the  **NavigationGroups** collection is a **[MailModule](Outlook.MailModule.md)** object.
    

## See also


[NavigationGroups Object](Outlook.NavigationGroups.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]