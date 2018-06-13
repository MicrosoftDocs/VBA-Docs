---
title: ContactsModule Object (Outlook)
keywords: vbaol11.chm3195
f1_keywords:
- vbaol11.chm3195
ms.prod: outlook
api_name:
- Outlook.ContactsModule
ms.assetid: fb183bd5-c72f-b38f-97e3-209a2a463d24
ms.date: 06/08/2017
---


# ContactsModule Object (Outlook)

Represents the  **Contacts** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **ContactsModule** object, derived from the **[NavigationModule](Outlook.NavigationModule.md)** object, provides access to the navigation groups contained in the **Contacts** navigation module of the Navigation Pane for an explorer. Use the **[GetNavigationModule](Outlook.NavigationModules.GetNavigationModule.md)** method or the **[Item](Outlook.NavigationModules.Item.md)** method of the **[Modules](Outlook.NavigationPane.Modules.md)** collection for the parent **[NavigationPane](Outlook.NavigationPane.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](Outlook.NavigationModule.NavigationModuleType.md)** property of the **NavigationModule** object to retrieve the navigation module type. If the **NavigationModuleType** property is set to **olModuleContacts**, you can then cast the **NavigationModule** object reference as a **ContactsModule** object to access the **[NavigationGroups](Outlook.ContactsModule.NavigationGroups.md)** property for that navigation module.

You can use the  **[Visible](Outlook.ContactsModule.Visible.md)** property to determine if the navigation module is visible and the **[Position](Outlook.ContactsModule.Position.md)** property to return or set the display position of the navigation module within the Navigation Pane. You can use the **[Name](Outlook.ContactsModule.Name.md)** property to return the display name of the **Contacts** navigation module within the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](Outlook.ContactsModule.Application.md)|
|[Class](Outlook.ContactsModule.Class.md)|
|[Name](Outlook.ContactsModule.Name.md)|
|[NavigationGroups](Outlook.ContactsModule.NavigationGroups.md)|
|[NavigationModuleType](Outlook.ContactsModule.NavigationModuleType.md)|
|[Parent](contactsmodule-parent-property-outlook.md)|
|[Position](Outlook.ContactsModule.Position.md)|
|[Session](contactsmodule-session-property-outlook.md)|
|[Visible](Outlook.ContactsModule.Visible.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
