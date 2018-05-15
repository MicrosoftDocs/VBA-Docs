---
title: NavigationGroups Object (Outlook)
keywords: vbaol11.chm3022
f1_keywords:
- vbaol11.chm3022
ms.prod: outlook
api_name:
- Outlook.NavigationGroups
ms.assetid: 07206203-36a9-7467-3a89-24fa2a7c2b1f
ms.date: 06/08/2017
---


# NavigationGroups Object (Outlook)

Contains a set of  **[NavigationGroup](Outlook.NavigationGroup.md)** objects that represent the navigation groups displayed by a navigation module in the Navigation Pane.


## Remarks

Use the  **[NavigationGroups](Outlook.MailModule.NavigationGroups.md)** property of the parent navigation module, such as a **[MailModule](Outlook.MailModule.md)** object, to return a **NavigationGroups** object.

Use the  **[Create](Outlook.NavigationGroups.Create.md)** method to create a new **NavigationGroup** object and add it to the collection. Use the **[Item](Outlook.NavigationGroups.Item.md)** method to retrieve a **NavigationGroup** object from the collection. Use the **[Delete](Outlook.NavigationGroups.Delete.md)** method of the **NavigationGroups** collection to create a new **NavigationGroup** object.

Use the  **[GetDefaultNavigationGroup](Outlook.NavigationGroups.GetDefaultNavigationGroup.md)** to return the default navigation group for a specified group type.


## Events



|**Name**|
|:-----|
|[NavigationFolderAdd](Outlook.NavigationGroups.NavigationFolderAdd.md)|
|[NavigationFolderRemove](Outlook.NavigationGroups.NavigationFolderRemove.md)|
|[SelectedChange](Outlook.NavigationGroups.SelectedChange.md)|

## Methods



|**Name**|
|:-----|
|[Create](Outlook.NavigationGroups.Create.md)|
|[Delete](Outlook.NavigationGroups.Delete.md)|
|[GetDefaultNavigationGroup](Outlook.NavigationGroups.GetDefaultNavigationGroup.md)|
|[Item](Outlook.NavigationGroups.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Outlook.NavigationGroups.Application.md)|
|[Class](navigationgroups-class-property-outlook.md)|
|[Count](navigationgroups-count-property-outlook.md)|
|[Parent](navigationgroups-parent-property-outlook.md)|
|[Session](navigationgroups-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
