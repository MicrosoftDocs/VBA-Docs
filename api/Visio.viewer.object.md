---
title: Viewer Object (Visio Viewer)
ms.prod: visio
ms.assetid: 4d25251a-5c4d-42d4-a73e-7e1e987ff593
ms.date: 06/08/2017
localization_priority: Normal
---


# Viewer Object (Visio Viewer)

The  **Viewer** object is a programmable ActiveX control that enables you to display Visio drawings (with limited functionality) on web pages and in Windows Forms, so that users who do not have Visio installed on their computers can view and interact with them.


## Remarks

With Visio Viewer, users can open, view, or print Visio drawings, even if they do not have Microsoft Visio 2013 installed. They cannot, however, edit or save drawings, or create a new Visio drawing. For that, they need to install Visio.

The  **Viewer** object is the entry point to the **Viewer** object model, and represents an instance of the Viewer control. The properties, events, and methods available in the **Viewer** object model let you load and unload Visio drawings in Visio Viewer, temporarily change properties and settings of the drawing, react to user input, and customize the Visio Viewer environment. In many cases, these members correspond to the options available to users in the Visio Viewer user interface (UI).

The following is a partial listing of the members of the  **Viewer** object and their functions and provides a sampling of the programming options available to developers. See the table of contents of this reference for the complete list of members. See [About Programming Visio Viewer](Visio.about.programming.md) for code samples that show how to get an instance of the **Viewer** object in the available development environments.

Use the  **[Load](Visio.Load.md)** method to load a Visio drawing into Visio Viewer, and use the **[Unload](Visio.Unload.md)** method to unload the drawing. You can also use the **[SRC](Visio.viewer.src.property.md)** property to get and set the file name and path for the current drawing.

Use the  **[DisplayAbout](Visio.DisplayAbout.md)**,  **[DisplayContextMenu](Visio.DisplayContextMenu.md)**,  **[DisplayHelp](Visio.DisplayHelp.md)**, and  **[DisplayPropertyDialog](Visio.DisplayPropertyDialog.md)** methods to display the dialog boxes and shortcut menus available in the Visio Viewer UI.

Use the  **[SelectShape](Visio.SelectShape.md)** method to select a particular shape in the drawing and the **[ShapeName](Visio.ShapeName.md)** and **[ShapeCount](Visio.ShapeCount.md)** properties to get information about shapes in the drawing.

Use properties such as  **[BackColor](Visio.viewer.backcolor.property.md)**,  **[GridVisible](Visio.GridVisible.md)**,  **[LayerColor](Visio.LayerColor.md)**,  **[PageColor](Visio.PageColor.md)**,  **[ScrollbarsVisible](Visio.ScrollbarsVisible.md)**, and  **[ToolbarVisible](Visio.ToolbarVisible.md)** to customize the appearance of the Visio Viewer UI.

Use the  **[CustomPropertyCount](Visio.CustomPropertyCount.md)**,  **[CustomPropertyName](Visio.CustomPropertyName.md)**, and  **[CustomPropertyValue](Visio.CustomPropertyValue.md)** properties to determine shape data (custom properties).

Use events such as  **[OnLayerChanged](Visio.OnLayerChanged.md)** and **[OnSelectionChanged](Visio.OnSelectionChanged.md)** to respond to user input.


