---
title: NavigationGroup Object (Outlook)
keywords: vbaol11.chm3199
f1_keywords:
- vbaol11.chm3199
ms.prod: outlook
api_name:
- Outlook.NavigationGroup
ms.assetid: a96eb2b1-af1f-71b2-6a0b-dcb5078beb1f
ms.date: 06/08/2017
---


# NavigationGroup Object (Outlook)

Represents a navigation group displayed by a navigation module in the Navigation Pane.


## Remarks

Use the  **[Item](Outlook.NavigationGroups.Item.md)** method to retrieve a **NavigationGroup** object from the **[NavigationGroups](Outlook.NavigationGroups.md)** collection of a parent navigation module, such as a **[MailModule](Outlook.MailModule.md)** object. Use the **[Create](Outlook.NavigationGroups.Create.md)** method of the **NavigationGroups** collection to create a new **NavigationGroup** object.

Use the  **[GroupType](Outlook.NavigationGroup.GroupType.md)** property to determine the group type of the navigation group and the **[Position](Outlook.NavigationGroup.Position.md)** property to return or set the display position of the navigation group within the Navigation Pane. You can also use the **[Name](Outlook.NavigationGroup.Name.md)** property to return or set the display name of the navigation group within the Navigation Pane.

Use the  **[NavigationFolders](Outlook.NavigationGroup.NavigationFolders.md)** property to return a **[NavigationFolders](Outlook.NavigationFolders.md)** object containing the navigation folders for the specified navigation group.


## Properties



|**Name**|
|:-----|
|[Application](Outlook.NavigationGroup.Application.md)|
|[Class](Outlook.NavigationGroup.Class.md)|
|[GroupType](Outlook.NavigationGroup.GroupType.md)|
|[Name](Outlook.NavigationGroup.Name.md)|
|[NavigationFolders](Outlook.NavigationGroup.NavigationFolders.md)|
|[Parent](navigationgroup-parent-property-outlook.md)|
|[Position](Outlook.NavigationGroup.Position.md)|
|[Session](navigationgroup-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
