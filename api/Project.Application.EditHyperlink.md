---
title: Application.EditHyperlink method (Project)
keywords: vbapj.chm1310
f1_keywords:
- vbapj.chm1310
ms.prod: project-server
api_name:
- Project.Application.EditHyperlink
ms.assetid: d652ccc4-207e-933f-c281-a2d5d7db0b76
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.EditHyperlink method (Project)

Edits the hyperlink of the selected assignment, resource, or task.


## Syntax

_expression_. `EditHyperlink`( `_Name_`, `_Address_`, `_SubAddress_`, `_ScreenTip_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the hyperlink as it appears in the Hyperlink field.|
| _Address_|Optional|**String**|The address of the target document.|
| _SubAddress_|Optional|**String**| A location within the target document.|
| _ScreenTip_|Optional|**String**|The ScreenTip text for the hyperlink.|

## Return value

 **Boolean**


## Remarks

Using the **EditHyperlink** method without specifying any arguments displays the **Edit Hyperlink** dialog box.


## Example

The following example first creates a hyperlink in the Gantt Chart view and then change the name to MyHyperLink.


```vb
Sub Edit_Hyperlink() 
 
 ViewApply Name:="&Gantt Chart" 
 SelectRow Row:=2, RowRelative:=False 
 InsertHyperlink Name:="https://MSDN", Address:="https://msdn.microsoft.com/", SubAddress:="", ScreenTip:="" 
 
 EditHyperlink Name:="MyHyperLink" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]