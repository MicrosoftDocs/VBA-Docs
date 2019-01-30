---
title: CommandBarPopup members (Office)
ms.prod: office
ms.assetid: 8ec16deb-bb74-2871-d837-f706c7a58f2b
ms.date: 01/30/2019
localization_priority: Normal
---


# CommandBarPopup members (Office)

Represents a pop-up control on a command bar.


## Methods

|Name|Description|
|:-----|:-----|
|[Copy](../../Office.CommandBarPopup.Copy.md)|Copies a command bar popup control to an existing command bar.|
|[Delete](../../Office.CommandBarPopup.Delete.md)|Deletes the **CommandBarPopup** object from its collection.|
|[Execute](../../Office.CommandBarPopup.Execute.md)|Runs the procedure or built-in command assigned to the specified **CommandBarPopup** control.|
|[Move](../../Office.CommandBarPopup.Move.md)|Moves the specified **CommandBarPopup** control to an existing command bar.|
|[Reset](../../Office.CommandBarPopup.Reset.md)|Resets a built-in **CommandBarPopup** control to its original function and face.|
|[SetFocus](../../Office.CommandBarPopup.SetFocus.md)|Moves the keyboard focus to the specified **CommandBarPopup** control. If the popup is disabled or isn't visible, this method will fail.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.CommandBarPopup.Application.md)|Gets an **Application** object that represents the container application for the **CommandBarPopup** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[BeginGroup](../../Office.CommandBarPopup.BeginGroup.md)|Gets **True** if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.|
|[BuiltIn](../../Office.CommandBarPopup.BuiltIn.md)|Is **True** if the specified command bar popup is a built-in command bar of the container application. Returns **False** if it's a custom command bar. Read-only.|
|[Caption](../../Office.CommandBarPopup.Caption.md)|Gets or sets the caption text for a command bar control. Read/write.|
|[CommandBar](../../Office.CommandBarPopup.CommandBar.md)|Gets a **CommandBar** object that represents the menu displayed by the specified pop-up control. Read-only.|
|[Controls](../../Office.CommandBarPopup.Controls.md)|Gets a **CommandBarControls** object that represents all the controls on a pop-up control. Read-only.|
|[Creator](../../Office.CommandBarPopup.Creator.md)|Gets a 32-bit integer that indicates the application in which the **CommandBarPopup** object was created. Read-only.|
|[DescriptionText](../../Office.CommandBarPopup.DescriptionText.md)|Gets or sets the description for a command barpopup control. Read/write.|
|[Enabled](../../Office.CommandBarPopup.Enabled.md)|Is **True** if the **CommandBarPopup** is enabled. Read/write.|
|[Height](../../Office.CommandBarPopup.Height.md)|Gets or sets the height of a **CommandBarPopup** control. Read/write.|
|[HelpContextId](../../Office.CommandBarPopup.HelpContextId.md)|Gets or sets the Help context Id number for the Help topic attached to the **CommandBarPopup** control. Read/write.|
|[HelpFile](../../Office.CommandBarPopup.HelpFile.md)|Gets or sets the file name for the Help topic attached to the **CommandBarPopup** control. Read/write.|
|[Id](../../Office.CommandBarPopup.Id.md)|Gets the ID for a built-in **CommandBarPopup** control. Read-only.|
|[Index](../../Office.CommandBarPopup.Index.md)|Gets a **Long** representing the index number for a **CommandBarPopup** object in the collection. Read-only.|
|[IsPriorityDropped](../../Office.CommandBarPopup.IsPriorityDropped.md)|Gets **True** if the **CommandBarPopup** control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the **Visible** property). Read-only.|
|[Left](../../Office.CommandBarPopup.Left.md)|Gets the horizontal position of the specified **CommandBarPopup** control (in pixels) relative to the left edge of the screen. Returns the distance from the left side of the docking area. Read-only.|
|[OLEMenuGroup](../../Office.CommandBarPopup.OLEMenuGroup.md)|Gets or sets a **MsoOLEMenuGroup** constant that represents the menu group that the specified command bar pop-up control belongs to when the menu groups of the OLE server are merged with the menu groups of an OLE client (that is, when an object of the container application type is embedded in another application). Read/write.|
|[OLEUsage](../../Office.CommandBarPopup.OLEUsage.md)|Gets or sets the OLE client and OLE server roles in which a **CommandBarPopup** control is used when two Microsoft Office applications are merged. Read/write.|
|[OnAction](../../Office.CommandBarPopup.OnAction.md)|Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a **CommandBarPopup** control. Read/write.|
|[Parameter](../../Office.CommandBarPopup.Parameter.md)|Gets or sets a string that an application can use to execute a command from a **CommandBarPopup** control. Read/write.|
|[Parent](../../Office.CommandBarPopup.Parent.md)|Gets the **Parent** object for the **CommandBarPopup** object. Read-only.|
|[Priority](../../Office.CommandBarPopup.Priority.md)|Gets or sets the priority of a **CommandBarPopup** control. Read/write.|
|[Tag](../../Office.CommandBarPopup.Tag.md)|Gets or sets information about the **CommandBarPopup** control, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.|
|[TooltipText](../../Office.CommandBarPopup.TooltipText.md)|Gets or sets the text displayed in a **CommandBarPopup's** **ScreenTip**. Read/write.|
|[Top](../../Office.CommandBarPopup.Top.md)|Gets the distance (in pixels) from the top edge of the specified **CommandBarPopup** control to the top edge of the screen. Read-only.|
|[Type](../../Office.CommandBarPopup.Type.md)|Gets the type of **CommandBarPopup** control. Read-only.|
|[Visible](../../Office.CommandBarPopup.Visible.md)|Gets or sets the **Visible** property of the **CommandBarPopup** control. Read/write.|
|[Width](../../Office.CommandBarPopup.Width.md)|Gets or sets the width (in pixels) of the specified **CommandBarPopup** control. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]