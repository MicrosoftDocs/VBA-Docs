---
title: CommandBar members (Office)
ms.prod: office
ms.assetid: e3756e7e-56a8-33a4-722f-640e5cc69b6d
ms.date: 01/30/2019
localization_priority: Normal
---


# CommandBar members (Office)

Represents a command bar in the container application. The **CommandBar** object is a member of the **CommandBars** collection.


## Methods

|Name|Description|
|:-----|:-----|
|[Delete](../../Office.CommandBar.Delete.md)|Deletes the **CommandBar** object from the collection.|
|[FindControl](../../Office.CommandBar.FindControl.md)|Gets a **CommandBarControl** object that fits a specified criteria.|
|[Reset](../../Office.CommandBar.Reset.md)|Resets a built-in command bar to its default configuration.|
|[ShowPopup](../../Office.CommandBar.ShowPopup.md)|Displays a command bar as a shortcut menu at the specified coordinates or at the current pointer coordinates.|


## Properties

|Name|Description|
|:-----|:-----|
|[AdaptiveMenu](../../Office.CommandBar.AdaptiveMenu.md)|Gets a **Boolean** value that specifies whether the command bar should include an adaptive menu. Read/write.|
|[Application](../../Office.CommandBar.Application.md)|Gets an **Application** object that represents the container application for the **CommandBar** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[BuiltIn](../../Office.CommandBar.BuiltIn.md)|Gets **True** if the specified command bar is a built-in command bar of the container application. Returns **False** if it is a custom command bar. Read-only.|
|[Context](../../Office.CommandBar.Context.md)|Gets or sets a string that determines where a command bar will be saved. The string is defined and interpreted by the application. Read/write.|
|[Controls](../../Office.CommandBar.Controls.md)|Gets a **CommandBarControls** object that represents all the controls on a command bar. Read-only.|
|[Creator](../../Office.CommandBar.Creator.md)|Gets a 32-bit integer that indicates the application in which the **CommandBar** object was created. Read-only.|
|[Enabled](../../Office.CommandBar.Enabled.md)|Gets or sets a **Boolean** value that specifies whether the specified **CommandBar** is enabled. Read/write.|
|[Height](../../Office.CommandBar.Height.md)|Gets or sets the height of a **CommandBar**. Read/write.|
|[Index](../../Office.CommandBar.Index.md)|Gets a **Long** representing the index number for a **CommandBar** object in the collection. Read-only.|
|[Left](../../Office.CommandBar.Left.md)|Gets or sets the horizontal distance (in pixels) of the **CommandBar** from the left edge of the object relative to the screen. Read/write.|
|[Name](../../Office.CommandBar.Name.md)|Gets the name of the built-in **CommandBar** object. Read-only.|
|[NameLocal](../../Office.CommandBar.NameLocal.md)|Gets the name of a built-in command bar as it's displayed in the language version of the container application, or returns or sets the name of a custom command bar. Read/write.|
|[Parent](../../Office.CommandBar.Parent.md)|Gets the **Parent** object for the **CommandBar** object. Read-only.|
|[Position](../../Office.CommandBar.Position.md)|Gets or sets an **MsoBarPosition** constant representing the position of a command bar. Read/write.|
|[Protection](../../Office.CommandBar.Protection.md)|Gets or sets an **MsoBarProtection** constant representing the way a command bar is protected from user customization. Read/write.|
|[RowIndex](../../Office.CommandBar.RowIndex.md)|Gets or sets the docking order of a command bar in relation to other command bars in the same docking area. Can be an integer greater than zero, or either of the following **MsoBarRow** constants: **msoBarRowFirst** or **msoBarRowLast**. Read/write.|
|[Top](../../Office.CommandBar.Top.md)|Gets or sets the distance from the top edge of the specified command bar, to the top edge of the screen. For docked command bars, this property returns or sets the distance from the command bar to the top of the docking area. Read/write.|
|[Type](../../Office.CommandBar.Type.md)|Gets the type of command bar. Read-only.|
|[Visible](../../Office.CommandBar.Visible.md)|Gets or sets the **Visible** property of the command bar. **True** if the command bar is visible. Read/write.|
|[Width](../../Office.CommandBar.Width.md)|Gets or sets the width (in pixels) of the specified command bar. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
