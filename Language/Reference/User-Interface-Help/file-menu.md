---
title: File menu
keywords: vbui6.chm2057556
f1_keywords:
- vbui6.chm2057556
ms.prod: office
ms.assetid: 1affc2fd-9a01-be54-de9c-bc51ddb6c417
ms.date: 11/24/2018
localization_priority: Normal
---


# File menu

> [!NOTE] 
> Some menu items are not available in all versions of the Visual Basic Editor.

## Close

|Command|Description|
|:------|:----------|
|**Close & Return to `<host application>`**|Closes the development environment and returns you to the host application. Visual Basic is hidden but remains in memory.|
|**Close Project**|Closes the current project. If the project contains any unsaved changes, you are prompted to save the project before closing.|

## Import File, Export File

Adds existing [modules](../../Glossary/vbe-glossary.md#module) and forms to your project, or saves a module or form as a separate file.

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Import File**|Displays the **[Import File](import-file-dialog-box.md)** dialog box, which allows you to add an existing module or form to the project. A copy of the file is added to the project, and the original file is left intact. If you import a form or module with the same name as an existing one, the file is added with a number appended to it.<br/><br/>Imported components appear in the Project Explorer window.|![Import File Toolbar button](../../../images/tbr_impt_ZA01201709.gif) | CTRL+M |
|**Export File**|Displays the **[Export File](export-file-dialog-box.md)** dialog box so that you can extract the active form or module from the project. The file is copied into an external file.<br/><br/>Not available if you have not selected a file in the Project Explorer.|![Export File Toolbar button](../../../images/tbr_expt_ZA01201702.gif) | CTRL+E |


## Make Project, New Project, Open Project

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Make `<Project>`**|Opens the **Make Project** dialog box so that you can build your project into a DLL.| | |
|**New Project** |Displays the **[New Project](new-project-dialog-box.md)** dialog box where you choose the type of [project](../../Glossary/vbe-glossary.md#project) that you want to create. If there is currently another project open when you create a new project, you will be prompted to save your work.<br/><br/>Available only at [design time](../../Glossary/vbe-glossary.md#design-time).| | |
|**Open Project** |Closes the current project or group project, if one is loaded, and opens an existing project or group of projects. You can open as many projects as your system resources permit.|![Open Project Toolbar button](../../../images/tbr_open_ZA01201720.gif) | CTRL+O |

## Print

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Print** |Prints forms and code to the printer specified in the Microsoft Windows **Control Panel**. See also the **[Print, Print Setup](print-setup-dialog-box.md)** dialog boxes. |![Print Toolbar button](../../../images/tbr_prnt_ZA01201725.gif) | CTRL+P|

## Remove Item

|Command|Description|
|:------|:----------|
|**Remove `<Item>`** |Permanently deletes the active [form](../../Glossary/vbe-glossary.md#form) or [module](../../Glossary/vbe-glossary.md#module) from the project. Not available when an item is not selected in the **Project Explorer**. <br/><br/>When removing a module from a project, make sure the remaining code doesn't refer to the removed item.<br/><br/>Before your item is deleted, you are asked if you want to export it as a file. If you click **Yes** in the message box, the **[Export File](export-file-dialog-box.md)** dialog box opens. If you click **No**, the item is deleted.<br/><br/>**IMPORTANT**: You cannot undo this action.|

## Save Host Document

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Save `<host document>`** |Saves the current project and all of its components (forms and [modules](../../Glossary/vbe-glossary.md#module)) with your document. A standalone project is saved separately as a project file with a .vba extension.<br/><br/>The **Save** command displays the **Save As** dialog box if this is the first time the project is being saved.|![Save Host Document Toolbar button](../../../images/tbr_save_ZA01201736.gif) | CTRL+S |


## See also

- [Menus and commands](../menus-commands.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]