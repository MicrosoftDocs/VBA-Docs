---
title: MailModule Object (Outlook)
keywords: vbaol11.chm3193
f1_keywords:
- vbaol11.chm3193
ms.prod: outlook
api_name:
- Outlook.MailModule
ms.assetid: df20efe5-be5c-952d-c6b7-20c20a83fda0
ms.date: 06/08/2017
---


# MailModule Object (Outlook)

Represents the  **Mail** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **MailModule** object, derived from the **[NavigationModule](Outlook.NavigationModule.md)** object, provides read-only access to the navigation groups contained in the **Mail** navigation module of the Navigation Pane for an explorer. Use the **[GetNavigationModule](Outlook.NavigationModules.GetNavigationModule.md)** method or the **[Item](Outlook.NavigationModules.Item.md)** method of the **[Modules](Outlook.NavigationPane.Modules.md)** collection for the parent **[NavigationPane](Outlook.NavigationPane.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](Outlook.NavigationModule.NavigationModuleType.md)** property of the **NavigationModule** object to retrieve the navigation module type. If the **NavigationModuleType** property is set to **olModuleMail**, you can then cast the **NavigationModule** object reference as a **MailModule** object to access the **[NavigationGroups](Outlook.MailModule.NavigationGroups.md)** property for that navigation module.


 **Note**  Unlike other navigation modules, such as the  **[CalendarModule](Outlook.CalendarModule.md)** object, you cannot create or delete navigation groups in the **MailModule** object.

You can use the  **[Visible](Outlook.MailModule.Visible.md)** property to determine if the navigation module is visible, and the **[Position](Outlook.MailModule.Position.md)** property to return or set the display position of the navigation module within the Navigation Pane. You can use the **[Name](Outlook.MailModule.Name.md)** property to return the display name of the **Mail** navigation module within the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](Outlook.MailModule.Application.md)|
|[Class](Outlook.MailModule.Class.md)|
|[Name](Outlook.MailModule.Name.md)|
|[NavigationGroups](Outlook.MailModule.NavigationGroups.md)|
|[NavigationModuleType](Outlook.MailModule.NavigationModuleType.md)|
|[Parent](mailmodule-parent-property-outlook.md)|
|[Position](Outlook.MailModule.Position.md)|
|[Session](mailmodule-session-property-outlook.md)|
|[Visible](Outlook.MailModule.Visible.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
