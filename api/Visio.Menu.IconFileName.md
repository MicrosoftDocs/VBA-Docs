---
title: Menu.IconFileName Method (Visio)
keywords: vis_sdr.chm13116350
f1_keywords:
- vis_sdr.chm13116350
ms.prod: visio
api_name:
- Visio.Menu.IconFileName
ms.assetid: 15becfe7-50fa-acdc-de51-9d041afd8aad
ms.date: 06/08/2017
localization_priority: Normal
---


# Menu.IconFileName Method (Visio)

Sets a custom icon file for a menu or toolbar item.


## Syntax

 _expression_. `IconFileName`( `_IconFileName_` )

 _expression_ A variable that represents a [Menu](./Visio.Menu.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _IconFileName_|Required| **String**|The path and name of the ICO, EXE, DLL, or VSL file to load.|

## Return value

Nothing


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The  **IconFileName** method loads the file that contains the icon, saves the bits, and discards the file name.

If the icon contains multiple images, Microsoft Visio chooses the best icon, based on both icon size and color depth.

Unless  _IconFileName_ is a fully qualified path, the application searches for the ICO, EXE, DLL, or VSL file in the folders indicated by the **Application** object's **AddonPaths** property (assuming that the **UIObject** object is in the Visio process).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how use the  **IconFileName** method to set the icon for a toolbar button. It retrieves a copy of the built-in Visio toolbars, adds a toolbar button, and sets the button icon.

Before running this code, replace  _path_ \ _filename_ with the full path to and name of a valid icon (.ico file) on your computer.

To restore the built-in Visio user interface after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
Public Sub IconFileName_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbarSet As Visio.ToolbarSet 
 Dim vsoToolbarItems As Visio.ToolbarItems 
 Dim vsoToolbarItem As Visio.ToolbarItem 
 
 'Get the UIObject object for the copy of the built-in toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 
 'Get the drawing window toolbar sets. 
 'NOTE: Use ItemAtID to get the toolbar set. 
 'Using vsoUIObject.ToolbarSets(visUIObjSetDrawing) will not work. 
 Set vsoToolbarSet = vsoUIObject.ToolbarSets.ItemAtID(visUIObjSetDrawing) 
 
 'Get the ToolbarItems collection. 
 Set vsoToolbarItems = vsoToolbarSet.Toolbars(0).ToolbarItems 
 
 'Add a new button in the first position. 
 Set vsoToolbarItem = vsoToolbarItems.AddAt(0) 
 
 'Set properties for the new toolbar button. 
 vsoToolbarItem.CntrlType = visCtrlTypeBUTTON 
 vsoToolbarItem.CmdNum = 1 
 
 vsoToolbarItem.IconFileName "path\filename " 
 
 'Use the new custom UI. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


