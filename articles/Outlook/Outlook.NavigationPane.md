---
title: NavigationPane Object (Outlook)
keywords: vbaol11.chm3021
f1_keywords:
- vbaol11.chm3021
ms.prod: outlook
api_name:
- Outlook.NavigationPane
ms.assetid: b6538c72-6115-99fc-c926-e0532a747823
ms.date: 06/08/2017
---


# NavigationPane Object (Outlook)

Represents the Navigation Pane displayed by the active  **[Explorer](Outlook.Explorer.md)** object.


## Remarks

Use the  **[NavigationPane](Outlook.Explorer.NavigationPane.md)** property of the **Explorer** object to retrieve a **NavigationPane** object, if one exists for the explorer.


 **Note**  Some  **Explorer** objects may not have a Navigation Pane.

Use the  **[IsCollapsed](Outlook.NavigationPane.IsCollapsed.md)** property to return or set the display mode of the Navigation Pane.

Use the  **[Modules](Outlook.NavigationPane.Modules.md)** property to return a **[NavigationModules](Outlook.NavigationModules.md)** object that represents the collection of navigation modules contained by the Navigation Pane. Use the **[DisplayedModuleCount](Outlook.NavigationPane.DisplayedModuleCount.md)** to return the count of **[NavigationModule](Outlook.NavigationModule.md)** objects currently displayed in the Navigation Pane and the **[CurrentModule](Outlook.NavigationPane.CurrentModule.md)** property to return or set the currently selected **NavigationModule** object.

Use the  **[ModuleSwitch](Outlook.NavigationPane.ModuleSwitch.md)** event to detect when the selected **NavigationModule** object changes in the Navigation Pane.


## Events



|**Name**|
|:-----|
|[ModuleSwitch](Outlook.NavigationPane.ModuleSwitch.md)|

## Properties



|**Name**|
|:-----|
|[Application](Outlook.NavigationPane.Application.md)|
|[Class](Outlook.NavigationPane.Class.md)|
|[CurrentModule](Outlook.NavigationPane.CurrentModule.md)|
|[DisplayedModuleCount](Outlook.NavigationPane.DisplayedModuleCount.md)|
|[IsCollapsed](Outlook.NavigationPane.IsCollapsed.md)|
|[Modules](Outlook.NavigationPane.Modules.md)|
|[Parent](navigationpane-parent-property-outlook.md)|
|[Session](navigationpane-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
