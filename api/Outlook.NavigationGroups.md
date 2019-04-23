---
title: NavigationGroups object (Outlook)
keywords: vbaol11.chm3022
f1_keywords:
- vbaol11.chm3022
ms.prod: outlook
api_name:
- Outlook.NavigationGroups
ms.assetid: 07206203-36a9-7467-3a89-24fa2a7c2b1f
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationGroups object (Outlook)

Contains a set of  **[NavigationGroup](Outlook.NavigationGroup.md)** objects that represent the navigation groups displayed by a navigation module in the navigation pane.


## Remarks

Use the  **[NavigationGroups](Outlook.MailModule.NavigationGroups.md)** property of the parent navigation module, such as a **[MailModule](Outlook.MailModule.md)** object, to return a **NavigationGroups** object.

Use the  **[Create](Outlook.NavigationGroups.Create.md)** method to create a new **NavigationGroup** object and add it to the collection. Use the **[Item](Outlook.NavigationGroups.Item.md)** method to retrieve a **NavigationGroup** object from the collection. Use the **[Delete](Outlook.NavigationGroups.Delete.md)** method of the **NavigationGroups** collection to create a new **NavigationGroup** object.

Use the  **[GetDefaultNavigationGroup](Outlook.NavigationGroups.GetDefaultNavigationGroup.md)** to return the default navigation group for a specified group type.


## Events



|Name|
|:-----|
|[NavigationFolderAdd](Outlook.NavigationGroups.NavigationFolderAdd.md)|
|[NavigationFolderRemove](Outlook.NavigationGroups.NavigationFolderRemove.md)|
|[SelectedChange](Outlook.NavigationGroups.SelectedChange.md)|

## Methods



|Name|
|:-----|
|[Create](Outlook.NavigationGroups.Create.md)|
|[Delete](Outlook.NavigationGroups.Delete.md)|
|[GetDefaultNavigationGroup](Outlook.NavigationGroups.GetDefaultNavigationGroup.md)|
|[Item](Outlook.NavigationGroups.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.NavigationGroups.Application.md)|
|[Class](Outlook.NavigationGroups.Class.md)|
|[Count](Outlook.NavigationGroups.Count.md)|
|[Parent](Outlook.NavigationGroups.Parent.md)|
|[Session](Outlook.NavigationGroups.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]