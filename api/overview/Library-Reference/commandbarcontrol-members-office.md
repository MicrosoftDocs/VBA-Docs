---
title: CommandBarControl members (Office)
ms.prod: office
ms.assetid: 1d2360e4-7511-a3a4-9959-2f7c8282bf99
ms.date: 01/30/2019
localization_priority: Normal
---


# CommandBarControl members (Office)

Represents a command bar control. The **CommandBarControl** object is a member of the **CommandBarControls** collection. The properties and methods of the **CommandBarControl** object are all shared by the **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** objects.


## Methods

|Name|Description|
|:-----|:-----|
|[Copy](../../Office.CommandBarControl.Copy.md)|Copies a command bar control to an existing command bar.|
|[Delete](../../Office.CommandBarControl.Delete.md)|Deletes the **CommandBarControl** object from its collection.|
|[Execute](../../Office.CommandBarControl.Execute.md)|Runs the procedure or built-in command assigned to the specified **CommandBarControl** control.|
|[Move](../../Office.CommandBarControl.Move.md)|Moves the specified **CommandBarControl** to an existing command bar.|
|[Reset](../../Office.CommandBarControl.Reset.md)|Resets a built-in command bar to its default configuration, or resets a built-in **CommandBarControl** to its original function and face.|
|[SetFocus](../../Office.CommandBarControl.SetFocus.md)|Moves the keyboard focus to the specified **CommandBarControl**. If the control is disabled or isn't visible, this method will fail.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.CommandBarControl.Application.md)|Gets an **Application** object that represents the container application for the **CommandBarControl** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[BeginGroup](../../Office.CommandBarControl.BeginGroup.md)|Gets **True** if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.|
|[BuiltIn](../../Office.CommandBarControl.BuiltIn.md)|Gets **True** if the specified command bar control is a built-in control of the container application. Returns **False** if it's a custom control, or if it's a built-in control whose **OnAction** property has been set. Read-only.|
|[Caption](../../Office.CommandBarControl.Caption.md)|Gets or sets the caption text for a command bar control. Read/write.|
|[Creator](../../Office.CommandBarControl.Creator.md)|Gets a 32-bit integer that indicates the application in which the **CommandBarControl** object was created. Read-only.|
|[DescriptionText](../../Office.CommandBarControl.DescriptionText.md)|Gets or sets the description for a command bar control. Read/write.|
|[Enabled](../../Office.CommandBarControl.Enabled.md)|Gets or sets a **Boolean** value specifying if the **CommandBarControl** is enabled. Read/write.|
|[Height](../../Office.CommandBarControl.Height.md)|Gets or sets the height of a **CommandBarControl** control. Read/write.|
|[HelpContextId](../../Office.CommandBarControl.HelpContextId.md)|Gets or sets the Help context Id number for the Help topic attached to the **CommandBarControl**. Read/write.|
|[HelpFile](../../Office.CommandBarControl.HelpFile.md)|Gets or sets the file name for the Help topic attached to the **CommandBarControl**. Read/write.|
|[Id](../../Office.CommandBarControl.Id.md)|Gets the ID for a built-in **CommandBarControl**. Read-only.|
|[Index](../../Office.CommandBarControl.Index.md)|Gets a **Long** representing the index number for a **CommandBarControl** object in the collection. Read-only.|
|[IsPriorityDropped](../../Office.CommandBarControl.IsPriorityDropped.md)|Gets **True** if the control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the **Visible** property). Read-only.|
|[Left](../../Office.CommandBarControl.Left.md)|Gets the horizontal position of the specified **CommandBarControl** (in pixels) relative to the left edge of the screen. Returns the distance from the left side of the docking area. Read-only.|
|[OLEUsage](../../Office.CommandBarControl.OLEUsage.md)|Gets or sets the OLE client and OLE server roles in which a **CommandBarControl** will be used when two Microsoft Office applications are merged. Read/write.|
|[OnAction](../../Office.CommandBarControl.OnAction.md)|Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a **CommandBarControl**. Read/write.|
|[Parameter](../../Office.CommandBarControl.Parameter.md)|Gets or sets a string that an application can use to execute a command from a **CommandBarControl**. Read/write.|
|[Parent](../../Office.CommandBarControl.Parent.md)|Gets the **Parent** object for the **CommandBarControl** object. Read-only.|
|[Priority](../../Office.CommandBarControl.Priority.md)|Gets or sets the priority of a **CommandBarControl**. A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Controls that can't fit in a single row drop off command bars from right to left. Read/write.|
|[Tag](../../Office.CommandBarControl.Tag.md)|Gets or sets information about the **CommandBarControl**, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.|
|[TooltipText](../../Office.CommandBarControl.TooltipText.md)|Gets or sets the text displayed in a **CommandBarControl's** **ScreenTip**. Read/write.|
|[Top](../../Office.CommandBarControl.Top.md)|Gets the distance (in pixels) from the top edge of the specified **CommandBarControl** to the top edge of the screen. Read-only.|
|[Type](../../Office.CommandBarControl.Type.md)|Gets the type of **CommandBarControl**. Read-only.|
|[Visible](../../Office.CommandBarControl.Visible.md)|Gets or sets the **Visible** property of the **CommandBarControl**. **True** if the **CommandBarControl** is visible. Read/write.|
|[Width](../../Office.CommandBarControl.Width.md)|Gets or sets the width (in pixels) of the specified **CommandBarControl**. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]