---
title: ToolbarItem.Enabled Property (Visio)
keywords: vis_sdr.chm13513455
f1_keywords:
- vis_sdr.chm13513455
ms.prod: visio
api_name:
- Visio.ToolbarItem.Enabled
ms.assetid: ef634f6d-ce50-4f88-c90f-75fbfaa02a29
ms.date: 06/08/2017
---


# ToolbarItem.Enabled Property (Visio)

Determines whether or not an object is currently enabled. Read/write.


## Syntax

 _expression_. `Enabled`

 _expression_ A variable that represents a [ToolbarItem](./Visio.ToolbarItem.md) object.


### Return value

Boolean


## Remarks




 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.


## Example

This example shows how to use the  **Enabled** property to enable hiding or showing a toolbar. The example adds a custom toolbar to the **Toolbars** collection. This toolbar appears in the Visio user interface and is available while the document is active.

To restore the built-in Visio toolbars after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Sub Enabled_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbars As Visio.Toolbars 
 Dim vsoToolbar As Visio.Toolbar 
 Dim vsoToolbarItem As Visio.ToolbarItem 
 
 'Check whether there are document custom toolbars. 
 If ThisDocument.CustomToolbars Is Nothing Then 
 
 'Check whether there are application custom toolbars. 
 If Visio.Application.CustomToolbars Is Nothing Then 
 
 'Use the built-in toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 
 Else 
 
 'Use the application custom toolbars. 
 Set vsoUIObject = Visio.Application.CustomToolbars.Clone 
 
 End If 
 
 Else 
 
 'Use the document custom toolbars. 
 Set vsoUIObject = ThisDocument.CustomToolbars 
 
 End If 
 
 'Get the Toolbars collection for the drawing window context. 
 Set vsoToolbars = vsoUIObject.ToolbarSets.ItemAtID( _ 
 Visio.visUIObjSetDrawing).Toolbars 
 
 'Add a toolbar to the collection. 
 Set vsoToolbar = vsoToolbars.Add 
 
 'Set the title of the toolbar. 
 vsoToolbar.Caption = "Example" 
 
 'Enable hiding or showing the toolbar. 
 vsoToolbar.Enabled = True 
 
 'Show the toolbar. 
 vsoToolbar.Visible = True 
 
 'Add an item to the toolbar. 
 Set vsoToolbarItem = vsoToolbar.ToolbarItems.Add 
 With vsoToolbarItem 
 
 'Set the new item to be a button. 
 .CntrlType = Visio.visCtrlTypeBUTTON 
 
 'Set the icon of the new button. 
 .FaceID = Visio.visIconIXCUSTOM_CARDS 
 
 'Set the CmdNum property of the new button. 
 .CmdNum = 1 
 
 'Set the Width property of the new button 
 'wide enough that the toolbar name is readable. 
 .Width = 100 
 
 End With 
 
 'Tell Visio to use the new UIOjbect object while 
 'this document is active. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


