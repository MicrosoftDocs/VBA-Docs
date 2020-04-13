---
title: MenuItems.AddAt method (Visio)
keywords: vis_sdr.chm13016015
f1_keywords:
- vis_sdr.chm13016015
ms.prod: visio
api_name:
- Visio.MenuItems.AddAt
ms.assetid: 1a987305-29f9-4f58-758b-a89d1e1911ca
ms.date: 06/08/2017
localization_priority: Normal
---


# MenuItems.AddAt method (Visio)

Creates a new **MenuItem** object at a specified index in the **MenuItems** collection.


## Syntax

_expression_. `AddAt`( `_lIndex_` )

_expression_ A variable that represents a **[MenuItems](Visio.MenuItems.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _lIndex_|Required| **Long**|The index at which to add the object.|

## Return value

MenuItem


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If the index is zero (0), the object is added at the beginning of the collection.

The beginning of a  **MenuItems** collection is the topmost menu item.


## Example

The following macro shows how to add a menu and menu item to the user interface.

This example assumes that you already have a macro called "MyMacro" in the Microsoft Visual Basic for Applications (VBA) project associated with the active Visio document.




```vb
Public Sub AddAt_Example() 
 Dim vsoUI As Visio.UIObject 
 Dim vsoMenuSets As Visio.MenuSets 
 Dim vsoMenuSet As Visio.MenuSet 
 Dim vsoMenus As Visio.Menus 
 Dim vsoMenu As Visio.Menu 
 Dim vsoMenuItems As Visio.MenuItems 
 Dim vsoMenuItem As Visio.MenuItem 
 
 'Get a UI object that represents Visio built-in menus. 
 Set vsoUI = Visio.Application.BuiltInMenus 
 
 'Get the MenuSets collection. 
 Set vsoMenuSets = vsoUI.MenuSets 
 
 'Get the drawing window menu set. 
 Set vsoMenuSet = vsoMenuSets.ItemAtID(visUIObjSetDrawing) 
 
 'Get the Menus collection. 
 Set vsoMenus = vsoMenuSet.Menus 
 
 'Add a Demo menu. 
 Set vsoMenu = vsoMenus.AddAt(7) 
 vsoMenu.Caption = "Demo" 
 
 'Get the MenuItems collection. 
 Set vsoMenuItems = vsoMenu.MenuItems 
 
 'Add a menu item to the new Demo menu. 
 Set vsoMenuItem = vsoMenuItems.Add 
 
 'Set the properties for the new menu item. 
 vsoMenuItem.Caption = "Run &MyMacro" 
 vsoMenuItem.AddOnName = "ThisDocument.MyMacro" 
 vsoMenuItem.ActionText = "Run MyMacro" 
 
 'Tell Visio to use the new UI when the document is active. 
 ThisDocument.SetCustomMenus vsoUI 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]