---
title: Set Visual Basic environment options (VBA)
keywords: vbhw6.chm1105240
f1_keywords:
- vbhw6.chm1105240
ms.prod: office
ms.assetid: ce85ae8c-9e02-2525-98e7-403d5a590d6c
ms.date: 12/27/2018
localization_priority: Normal
---


# Set Visual Basic environment options

You can set the behavior and look of the Visual Basic development environment by using the **[Options](../reference/user-interface-help/options-dialog-box.md)** dialog box. Use the:

- **Editor** tab to specify Code window and Project window settings.   
- **Editor Format** tab to specify the appearance of your code.   
- **General** tab to specify form, error handling, and compile settings for your project.   
- **Docking** tab to specify whether a window is attached or "anchored" to one edge of other dockable or application windows.
    
To set environment options, on the **[Tools](../reference/user-interface-help/tools-menu.md)** menu of the Visual Basic editor, choose **Options**. Each option is described in the following sections.
    
## Editor

|Option|Description|
|:-----|:-----|
|**Auto Syntax Check**|Visual Basic automatically verifies correct syntax after you enter a line of code.|
|**Require Variable Declaration**|Explicit variable declarations are required in [modules](../Glossary/vbe-glossary.md#module).|
|**Auto Indent**|After tabbing the first line of code, all subsequent lines start at that tab location.|
|**Tab Width**|The tab width, which can range from 1&ndash;32 spaces. (Default is 4 spaces.)|
|**Default to Full Module View**|[Procedures](../Glossary/vbe-glossary.md#procedure) for new modules are displayed in the Code window as a single, scrollable list or one procedure at a time.|
|**Procedure Separator**|Display separator bars at the end of each procedure in the Code window.|
|**Auto List Members**|At the insertion point, Visual Basic displays information that logically completes a statement.|
|**Auto Quick Info**|Information about functions and their [arguments](../Glossary/vbe-glossary.md#argument) is displayed as you type.|
|**Auto Data Tips**|Automatically display the value of any [variable](../Glossary/vbe-glossary.md#variable) on which you place the mouse pointer. Available only in [break mode](../Glossary/vbe-glossary.md#break-mode).|
|**Drag-Drop in Text Editing**|Code elements can be dragged from the Code window into the Immediate or Watch windows.|



## Editor Format

|Option|Description|
|:-----|:-----|
|**Foreground**, **Background**, and **Indicator**|The color of different categories of text listed in the **Code Colors** list.|
|**Font**|The font used for displaying code.|
|**Size**|The size of the font used for code.|
|**Margin Indicator Bar**|Display the **Margin Indicator Bar**.|



## General

|Option|Description|
|:-----|:-----|
|**Show Grid**|Display a grid on a form.|
|**Grid Units**|Lists the unit of measurement for units in the grid.|
|**Width**|The width of the grid cells on a form.|
|**Height**|The height of the grid cells on a form.|
|**Align Controls to Grid**|Automatically position the outer edge of controls on the closest grid lines.|
|**Show ToolTips**|Display ToolTips for toolbar buttons.|
|**Collapse Proj. Hides Windows**|Automatically close the project, **UserForm**, object, or module windows when a [project](../Glossary/vbe-glossary.md#project) is collapsed in the **Project Explorer**.|
|**Notify Before State Loss**|Display a message that a requested action will cause all module-level variables to be reset for a running project.|
|**Break on All Errors**|Any error causes the project to enter break mode, whether or not an error handler is active, and whether or not the code is in a [class module](../Glossary/vbe-glossary.md#class-module).|
|**Break in Class Module**|Any unhandled error produced in a class module causes the project to enter break mode at the line of code which produced the error.|
|**Break on Unhandled Errors**|Any other unhandled error causes the project to enter break mode.|
|**Compile On Demand**|A project is fully compiled before it starts, or code is compiled as needed.|
|**Background Compile**|Use idle time during run time to finish compiling the project in the background (available only if **Compile On Demand** is set).|



## Docking

|Option|Description|
|:-----|:-----|
|The check box for the appropriate window|A window can be anchored to an adjacent dockable window or the Visual Basic Editor window.|

## See also

- [Visual Basic how-to topics](../reference/user-interface-help/visual-basic-how-to-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
