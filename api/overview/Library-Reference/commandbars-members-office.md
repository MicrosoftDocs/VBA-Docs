---
title: CommandBars members (Office)
ms.prod: office
ms.assetid: c11db22d-b7bb-20a2-a455-e441cb8d5bc0
ms.date: 01/30/2019
localization_priority: Normal
---


# CommandBars members (Office)

A collection of **CommandBar** objects that represent the command bars in the container application.


## Events

|Name|Description|
|:-----|:-----|
|[OnUpdate](../../Office.CommandBars.OnUpdate.md)|Occurs when any change is made to a command bar.|


## Methods

|Name|Description|
|:-----|:-----|
|[Add](../../Office.CommandBars.Add.md)|Creates a new command bar and adds it to the collection of command bars.|
|[CommitRenderingTransaction](../../Office.CommandBars.CommitRenderingTransaction.md)|Commits the rendering transaction. Returns **Nothing**.|
|[ExecuteMso](../../Office.CommandBars.ExecuteMso.md)|Executes the control identified by the _idMso_ parameter.|
|[FindControl](../../Office.CommandBars.FindControl.md)|Gets a **CommandBarControl** object that fits a specified criteria.|
|[FindControls](../../Office.CommandBars.FindControls.md)|Gets the **CommandBarControls** collection that fits the specified criteria.|
|[GetEnabledMso](../../Office.CommandBars.GetEnabledMso.md)|Returns **True** if the control identified by the _idMso_ parameter is enabled.|
|[GetImageMso](../../Office.CommandBars.GetImageMso.md)|Returns an **IPictureDisp** object of the control image identified by the _idMso_ parameter scaled to the dimensions specified by width and height.|
|[GetLabelMso](../../Office.CommandBars.GetLabelMso.md)|Returns the label of the control identified by the _idMso_ parameter as a **String**.|
|[GetPressedMso](../../Office.CommandBars.GetPressedMso.md)|Returns a value indicating whether the **toggleButton** control identified by the _idMso_ parameter is pressed.|
|[GetScreentipMso](../../Office.CommandBars.GetScreentipMso.md)|Returns the screentip of the control identified by the _idMso_ parameter as a **String**.|
|[GetSupertipMso](../../Office.CommandBars.GetSupertipMso.md)|Returns the supertip of the control identified by the _idMso_ parameter as a **String**.|
|[GetVisibleMso](../../Office.CommandBars.GetVisibleMso.md)|Returns **True** if the control identified by the _idMso_ parameter is visible.|
|[ReleaseFocus](../../Office.CommandBars.ReleaseFocus.md)|Releases the user interface focus from all command bars.|


## Properties

|Name|Description|
|:-----|:-----|
|[ActionControl](../../Office.CommandBars.ActionControl.md)|Gets the **CommandBarControl** object whose **OnAction** property is set to the running procedure. Read-only.|
|[ActiveMenuBar](../../Office.CommandBars.ActiveMenuBar.md)|Gets a **CommandBar** object that represents the active menu bar in the container application. Read-only.|
|[AdaptiveMenus](../../Office.CommandBars.AdaptiveMenus.md)|Checks or unchecks the check box control for the option to show menus in Microsoft Office as full or personalized. Read/write.|
|[Application](../../Office.CommandBars.Application.md)|Gets an **Application** object that represents the container application for the **CommandBars** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Count](../../Office.CommandBars.Count.md)|Gets a count of the number of command bars in the host application. Read-only.|
|[Creator](../../Office.CommandBars.Creator.md)|Gets a 32-bit integer that indicates the application in which the **CommandBars** object was created. Read-only.|
|[DisableAskAQuestionDropdown](../../Office.CommandBars.DisableAskAQuestionDropdown.md)|Is **True** if the **Answer Wizard** dropdown menu is enabled. Read/write.|
|[DisableCustomize](../../Office.CommandBars.DisableCustomize.md)|Is **True** if toolbar customization is disabled. Read/write.|
|[DisplayFonts](../../Office.CommandBars.DisplayFonts.md)|Is **True** if the font names in the **Font** box are displayed in their actual fonts. Read/write.|
|[DisplayKeysInTooltips](../../Office.CommandBars.DisplayKeysInTooltips.md)|Is **True** if shortcut keys are displayed in the **ToolTips** for each command bar control. Read/write.|
|[DisplayTooltips](../../Office.CommandBars.DisplayTooltips.md)|Is **True** if ScreenTips are displayed whenever the user positions the pointer over command bar controls. Read/write.|
|[Item](../../Office.CommandBars.Item.md)|Gets a **CommandBar** object from the **CommandBars** collection. Read-only.|
|[LargeButtons](../../Office.CommandBars.LargeButtons.md)|Is **True** if the toolbar buttons displayed are larger than normal size. Read/write.|
|[MenuAnimationStyle](../../Office.CommandBars.MenuAnimationStyle.md)|Gets or sets an **MsoMenuAnimation** that represents the way a command bar is animated. Read/write.|
|[Parent](../../Office.CommandBars.Parent.md)|Gets the **Parent** object for the **CommandBars** object. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
