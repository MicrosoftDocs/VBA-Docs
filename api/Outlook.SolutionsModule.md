---
title: SolutionsModule object (Outlook)
keywords: vbaol11.chm3371
f1_keywords:
- vbaol11.chm3371
ms.prod: outlook
api_name:
- Outlook.SolutionsModule
ms.assetid: 4597765e-a95d-bf07-2ac4-103218ebc696
ms.date: 06/08/2017
localization_priority: Normal
---


# SolutionsModule object (Outlook)

Represents the  **Solutions** navigation module in the navigation pane of an explorer.


## Remarks

The  **Solutions** navigation module contains folders that developers of individual add-ins want to expose to users in the navigation pane. Each solution has one root folder under the **Solutions** module, and each root folder can contain subfolders that hold heterogeneous Outlook items.

To add solution folders programmatically to the  **Solutions** module, use the **SolutionsModule** object, which is derived from the **[NavigationModule](Outlook.NavigationModule.md)** object.

To obtain an object for the  **Solutions** module, you must first determine whether the **Solutions** module exists in the navigation pane. To do that, use the **Modules** property for the **[NavigationPane](Outlook.NavigationPane.md)** object to obtain a **[NavigationModules](Outlook.NavigationModules.md)** collection, and then specify the argument **olModuleSolutions** in the **[GetNavigationModule](Outlook.NavigationModules.GetNavigationModule.md)** method of the **NavigationModules** collection.

If the call is successful, you can then cast the returned  **NavigationModule** object reference as a **SolutionsModule** object to access the properties and methods for that navigation module.

To add a solution root folder and its subfolders, pass a  **[Folder](Outlook.Folder.md)** object reference to the **[AddSolution](Outlook.SolutionsModule.AddSolution.md)** method of the **SolutionsModule** object. The default position of the **Solutions** module on the navigation pane is '9'.

If no solutions have been added to the  **Solutions** module, it is not visible in the navigation pane, and any attempt to set the **[Position](Outlook.SolutionsModule.Position.md)** or the **[Visible](Outlook.SolutionsModule.Visible.md)** properties of the **SolutionsModule** object raises an error. In addition, any attempt to set the **SolutionsModule** as the **[CurrentModule](Outlook.NavigationPane.CurrentModule.md)** property of the **NavigationPane** object raises an error.


## Example

To see an example of an add-in that adds folders to the  **Solutions** module, see the article[Programming the Outlook 2010 Solutions Module](https://msdn.microsoft.com/library/ee692173%28office.14%29.aspx) on MSDN. The add-in in the article renames the **Solutions** module as **Solution Demo**, adds calendar, contacts, and tasks folders as subfolders to the solution root folder, sets custom icons for each of the subfolders, and customizes the navigation pane to move and enlarge the button for the  **Solution Demo** module.


## Methods



|Name|
|:-----|
|[AddSolution](Outlook.SolutionsModule.AddSolution.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.SolutionsModule.Application.md)|
|[Class](Outlook.SolutionsModule.Class.md)|
|[Name](Outlook.SolutionsModule.Name.md)|
|[NavigationModuleType](Outlook.SolutionsModule.NavigationModuleType.md)|
|[Parent](Outlook.SolutionsModule.Parent.md)|
|[Position](Outlook.SolutionsModule.Position.md)|
|[Session](Outlook.SolutionsModule.Session.md)|
|[Visible](Outlook.SolutionsModule.Visible.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]