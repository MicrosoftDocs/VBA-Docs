---
title: Application.WindowUnhide method (Project)
keywords: vbapj.chm704
f1_keywords:
- vbapj.chm704
ms.prod: project-server
api_name:
- Project.Application.WindowUnhide
ms.assetid: 438693a7-5b99-e373-6d28-9a42dfcda7d1
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowUnhide method (Project)

Shows a hidden window.


## Syntax

_expression_. `WindowUnhide`( `_Name_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of a hidden window to show. The name of a window is the exact text that appears in the title bar of the window. If Name is omitted, the  **Unhide** dialog box appears, which prompts the user to show a hidden window in the active project.|

## Return value

 **Boolean**


## Example

The following example unhides all open windows.


```vb
Sub UnhideAllWindows() 
 
 Dim I As Long ' Index for For...Next loop 
 
 For I = 1 To Windows.Count 
 If Not Windows(I).Visible Then 
 
 WindowUnhide Windows(I).Caption 
 End If 
 Next I 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]