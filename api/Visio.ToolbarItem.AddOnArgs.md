---
title: ToolbarItem.AddOnArgs property (Visio)
keywords: vis_sdr.chm13513045
f1_keywords:
- vis_sdr.chm13513045
ms.prod: visio
api_name:
- Visio.ToolbarItem.AddOnArgs
ms.assetid: 2e498e9e-021c-3857-2420-f7f8cc5de6c5
ms.date: 06/08/2017
localization_priority: Normal
---


# ToolbarItem.AddOnArgs property (Visio)

Gets or sets the argument string that you send to the add-on associated with a particular toolbar. Read/write.


## Syntax

_expression_.**AddOnArgs**

_expression_ A variable that represents a **[ToolbarItem](Visio.ToolbarItem.md)** object.


## Return value

String


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

An argument's string can be anything appropriate for the add-on. However, the arguments are packaged together with other information into a command string, which cannot exceed 127 characters. For best results, limit arguments to 50 characters.

An object's  **AddOnName** property indicates the name of the add-on to which the arguments are sent.

 Beginning with Visio 2002, the **AddOnName** property used in the following example cannot execute a string that contains arbitrary Microsoft Visual Basic code. To call code that in previous versions of Visio you would have passed to the **AddOnName** property, move it to a procedure in a document's Visual Basic project that is called from the **AddOnName** property, as shown in the following example.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]