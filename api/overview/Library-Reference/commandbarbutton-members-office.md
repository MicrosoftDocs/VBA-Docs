---
title: CommandBarButton members (Office)
ms.prod: office
ms.assetid: 69fe57fe-dabc-9379-283c-d0a51a775592
ms.date: 01/30/2019
localization_priority: Normal
---


# CommandBarButton members (Office)

Represents a button control on a command bar.


## Events

|Name|Description|
|:-----|:-----|
|[Click](../../Office.CommandBarButton.Click.md)|Occurs when the user clicks a **CommandBarButton** object.|


## Methods

|Name|Description|
|:-----|:-----|
|[Copy](../../Office.CommandBarButton.Copy.md)|Copies a command bar button control to an existing command bar.|
|[CopyFace](../../Office.CommandBarButton.CopyFace.md)|Copies the face of a command bar button control to the Clipboard.|
|[Delete](../../Office.CommandBarButton.Delete.md)|Deletes the **CommandBarButton** object from its collection.|
|[Execute](../../Office.CommandBarButton.Execute.md)|Runs the procedure or built-in command assigned to the specified **CommandBarButton** control.|
|[Move](../../Office.CommandBarButton.Move.md)|Moves the specified **CommandBarButton** control to an existing command bar.|
|[PasteFace](../../Office.CommandBarButton.PasteFace.md)|Pastes the contents of the Clipboard onto a **CommandBarButton**.|
|[Reset](../../Office.CommandBarButton.Reset.md)|Resets a built-in **CommandBarButton** control to its original function and face.|
|[SetFocus](../../Office.CommandBarButton.SetFocus.md)|Moves the keyboard focus to the specified **CommandBarButton** control. If the button is disabled or isn't visible, this method will fail.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.CommandBarButton.Application.md)|Gets an **Application** object that represents the container application for the **CommandBarButton** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[BeginGroup](../../Office.CommandBarButton.BeginGroup.md)|Gets **True** if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.|
|[BuiltIn](../../Office.CommandBarButton.BuiltIn.md)|Is **True** if the specified command bar control is a control of the container application. Returns **False** if it's a custom control, or if it's a built-in control whose **OnAction** property has been set. Read-only.|
|[BuiltInFace](../../Office.CommandBarButton.BuiltInFace.md)|Is **True** if the face of a command bar button control is its original built-in face. Read/write.|
|[Caption](../../Office.CommandBarButton.Caption.md)|Gets or sets the caption text for a command bar control. Read/write.|
|[Creator](../../Office.CommandBarButton.Creator.md)|Gets a 32-bit integer that indicates the application in which the **CommandBarButton** object was created. Read-only.|
|[DescriptionText](../../Office.CommandBarButton.DescriptionText.md)|Gets or sets the description for a command bar button control. Read/write.|
|[Enabled](../../Office.CommandBarButton.Enabled.md)|**True** if the specified **CommandBar** or **CommandBarControl** is enabled. Read/write .|
|[FaceId](../../Office.CommandBarButton.FaceId.md)|Gets or sets the Id number for the face of a **CommandBarButton** control. Read/write.|
|[Height](../../Office.CommandBarButton.Height.md)|Gets or sets the height of a command bar control. Read/write.|
|[HelpContextId](../../Office.CommandBarButton.HelpContextId.md)|Gets or sets the Help context Id number for the Help topic attached to the **CommandBarButton** control. Read/write.|
|[HelpFile](../../Office.CommandBarButton.HelpFile.md)|Gets or sets the file name for the Help topic attached to the **CommandBarButton** control. Read/write.|
|[HyperlinkType](../../Office.CommandBarButton.HyperlinkType.md)|Gets or sets an **MsoCommandBarButtonHyperlinkType** constant that represents the type of hyperlink associated with the specified command bar button. Read/write.|
|[Id](../../Office.CommandBarButton.Id.md)|Gets the ID for a built-in **CommandBarButton** control. Read-only.|
|[Index](../../Office.CommandBarButton.Index.md)|Gets a **Long** representing the index number for a **CommandBarButton** object in the collection. Read-only.|
|[IsPriorityDropped](../../Office.CommandBarButton.IsPriorityDropped.md)|Gets **True** if the **CommandBarButton** control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the **Visible** property). Read-only.|
|[Left](../../Office.CommandBarButton.Left.md)|Gets or sets the horizontal position of the specified **CommandBarButton** control (in pixels) relative to the left edge of the screen. Returns the distance from the left side of the docking area. Read-only.|
|[Mask](../../Office.CommandBarButton.Mask.md)|Gets or sets an **IPictureDisp** object representing the mask image of a **CommandBarButton** object. The mask image determines what parts of the button image are transparent. Read/write.|
|[OLEUsage](../../Office.CommandBarButton.OLEUsage.md)|Gets or sets the OLE client and OLE server roles in which a **CommandBarButton** control will be used when two Microsoft Office applications are merged. Read/write.|
|[OnAction](../../Office.CommandBarButton.OnAction.md)|Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a **CommandBarButton** control. Read/write.|
|[Parameter](../../Office.CommandBarButton.Parameter.md)|Gets or sets a string that an application can use to execute a command from a **CommandBarButton** control. Read/write.|
|[Parent](../../Office.CommandBarButton.Parent.md)|Gets the **Parent** object for the **CommandBarButton** object. Read-only.|
|[Picture](../../Office.CommandBarButton.Picture.md)|Gets or sets an **IPictureDisp** object representing the image of a **CommandBarButton** object. Read/write.|
|[Priority](../../Office.CommandBarButton.Priority.md)|Gets or sets the priority of a **CommandBarButton** control. A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Controls that can't fit in a single row drop off command bars from right to left. Read/write.|
|[ShortcutText](../../Office.CommandBarButton.ShortcutText.md)|Gets or sets the shortcut key text displayed next to a **CommandBarButton** control when the button appears on a menu, submenu, or shortcut menu. Read/write.|
|[State](../../Office.CommandBarButton.State.md)|Gets or sets the appearance of a **CommandBarButton** control. Read/write.|
|[Style](../../Office.CommandBarButton.Style.md)|Gets or sets the way a **CommandBarButton** control is displayed. Read/write.|
|[Tag](../../Office.CommandBarButton.Tag.md)|Gets or sets information about the **CommandBarButton** control, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.|
|[TooltipText](../../Office.CommandBarButton.TooltipText.md)|Gets or sets the text displayed in a **CommandBarButton's** **ScreenTip**. Read/write.|
|[Top](../../Office.CommandBarButton.Top.md)|Gets the distance (in pixels) from the top edge of the specified **CommandBarButton** control to the top edge of the screen. Read-only.|
|[Type](../../Office.CommandBarButton.Type.md)|Gets the type of **CommandBarButton** control. Read-only.|
|[Visible](../../Office.CommandBarButton.Visible.md)|Gets or sets the **Visible** property of the **CommandBarButton** control. **True** if the **CommandBarButton** is visible. Read/write.|
|[Width](../../Office.CommandBarButton.Width.md)|Gets or sets the width (in pixels) of the specified **CommandBarButton** control. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]