---
title: FileDialog members (Office)
description: Provides file dialog box functionality similar to the functionality of the standard Open and Save dialog boxes found in Microsoft Office applications.
ms.prod: office
ms.assetid: b6b7e87e-9420-0649-2feb-6d8f36bb53bc
ms.date: 01/30/2019
localization_priority: Normal
---


# FileDialog members (Office)

Provides file dialog box functionality similar to the functionality of the standard **Open** and **Save** dialog boxes found in Microsoft Office applications.


## Methods

|Name|Description|
|:-----|:-----|
|[Execute](../../Office.FileDialog.Execute.md)|Carries out a user's action right after the **Show** method is invoked.|
|[Show](../../Office.FileDialog.Show.md)|Displays a file dialog box and returns a **Long** indicating whether the user pressed the **Action** button (-1) or the **Cancel** button (0). When you call the **Show** method, no more code executes until the user dismisses the file dialog box. In the case of **Open** and **SaveAs** dialog boxes, use the **Execute** method right after the **Show** method to carry out the user's action.|

## Properties

|Name|Description|
|:-----|:-----|
|[AllowMultiSelect](../../Office.FileDialog.AllowMultiSelect.md)|Is **True** if the user is allowed to select multiple files from a file dialog box. Read/write.|
|[Application](../../Office.FileDialog.Application.md)|Gets an **Application** object that represents the container application for the **FileDialog** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[ButtonName](../../Office.FileDialog.ButtonName.md)|Gets or sets a **String** representing the text that is displayed on the action button of a file dialog box. Read/write.|
|[Creator](../../Office.FileDialog.Creator.md)|Gets a 32-bit integer that indicates the application in which the **FileDialog** object was created. Read-only.|
|[DialogType](../../Office.FileDialog.DialogType.md)|Gets an **MsoFileDialogType** constant representing the type of file dialog box that the **FileDialog** object is set to display. Read-only.|
|[FilterIndex](../../Office.FileDialog.FilterIndex.md)|Gets or sets a **Long** indicating the default file filter of a file dialog box. The default filter determines which types of files are displayed when the file dialog box is first opened. Read/write.|
|[Filters](../../Office.FileDialog.Filters.md)|Gets a **FileDialogFilters** collection. Read-only.|
|[InitialFileName](../../Office.FileDialog.InitialFileName.md)|Sets or returns a **String** representing the path or file name that is initially displayed in a file dialog box. Read/write.|
|[InitialView](../../Office.FileDialog.InitialView.md)|Gets or sets an **MsoFileDialogView** constant representing the initial presentation of files and folders in a file dialog box. Read/write.|
|[Item](../../Office.FileDialog.Item.md)|Gets the text associated with an object. Read-only.|
|[Parent](../../Office.FileDialog.Parent.md)|Gets the **Parent** object for the **FileDialog** object. Read-only.|
|[SelectedItems](../../Office.FileDialog.SelectedItems.md)|Gets a **FileDialogSelectedItems** collection. This collection contains a list of the paths of the files that a user selected from a file dialog box displayed by using the **Show** method of the **FileDialog** object. Read-only.|
|[Title](../../Office.FileDialog.Title.md)|Gets or sets the title of a file dialog box displayed by using the **FileDialog** object. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
