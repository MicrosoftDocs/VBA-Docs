---
title: Application.EditClearHyperlink method (Project)
keywords: vbapj.chm1316
f1_keywords:
- vbapj.chm1316
ms.prod: project-server
api_name:
- Project.Application.EditClearHyperlink
ms.assetid: 386e9e73-5c65-0baf-2125-4dbb50675eb1
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.EditClearHyperlink method (Project)

Clears the Hyperlink, Hyperlink Address, Hyperlink SubAddress, and HyperlinkHREF fields of the selected assignment, resource, or task.


## Syntax

_expression_. `EditClearHyperlink`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Example

The following example first creates a hyperlink in the Gantt Chart view and then clears it.


```vb
Sub EditClear_Hyperlink() 
 
 'Activate Gantt Chart view 
 ViewApply Name:="&Gantt Chart" 
 SelectRow Row:=2, RowRelative:=False 
 InsertHyperlink Name:="https://MSDN", Address:="https://msdn.microsoft.com/", SubAddress:="", ScreenTip:="" 
 
 EditClearHyperlink 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]