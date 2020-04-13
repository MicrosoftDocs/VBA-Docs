---
title: Explorers.Add method (Outlook)
keywords: vbaol11.chm122
f1_keywords:
- vbaol11.chm122
ms.prod: outlook
api_name:
- Outlook.Explorers.Add
ms.assetid: c3db3c6f-6441-c23e-06f2-afb5b61e5662
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorers.Add method (Outlook)

Creates a new instance of the explorer window.


## Syntax

_expression_.**Add** (_Folder_, _DisplayMode_)

_expression_ A variable that represents an [Explorers](Outlook.Explorers.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Folder_|Required| **Variant**|The **Variant** object to display in the explorer window when it is created.|
| _DisplayMode_|Optional| **Long**|The display mode of the folder. Can be one of the  **[OlFolderDisplayMode](Outlook.OlFolderDisplayMode.md)** constants.|

## Return value

An **[Explorer](Outlook.Explorer.md)** object that represents a new instance of the window.


## Remarks

The  _Folder_ argument can represent either a **[Folder](Outlook.Folder.md)** object or the URL to that folder.

The explorer window is initially hidden. You must call the  **[Display](Outlook.Explorer.Display.md)** method of the **Explorer** object to make it visible.


## Example

The following VBA example displays the Drafts folder in an explorer window without a Navigation Pane or Folder List.


```vb
Sub DisplayDrafts() 
 
 Dim myExplorers As Outlook.Explorers 
 
 Dim myOlExpl As Outlook.Explorer 
 
 Dim myFolder As Outlook.Folder 
 
 
 
 Set myExplorers = Application.Explorers 
 
 Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder _ 
 
 (olFolderDrafts) 
 
 Set myOlExpl = myExplorers.Add _ 
 
 (myFolder, olFolderDisplayNoNavigation) 
 
 myOlExpl.Display 
 
End Sub
```


## See also


[Explorers Object](Outlook.Explorers.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]