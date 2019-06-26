---
title: Application.InsertHyperlink method (Project)
keywords: vbapj.chm1309
f1_keywords:
- vbapj.chm1309
ms.prod: project-server
api_name:
- Project.Application.InsertHyperlink
ms.assetid: d5a6ffc3-8cfe-e6c9-c347-4e3a739f6b1a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.InsertHyperlink method (Project)

Inserts a hyperlink on the selected assignment, resource, or task.


## Syntax

_expression_. `InsertHyperlink`( `_Name_`, `_Address_`, `_SubAddress_`, `_ScreenTip_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the hyperlink as it appears in the Hyperlink field.|
| _Address_|Optional|**String**|The address of the target document.|
| _SubAddress_|Optional|**String**|A location within the target document.|
| _ScreenTip_|Optional|**String**|The ScreenTip text for the hyperlink.|

## Return value

 **Boolean**


## Remarks

Using the  **InsertHyperlink** method without specifying any arguments displays the **Insert Hyperlink** dialog box.


## Example

The following example inserts a hyperlink in a Gantt Chart view.


```vb
Sub Insert_Hyperlink() 
 
 'Activate Gantt Chart view 
 ViewApply Name:="&Gantt Chart" 
 
 SelectRow Row:=2, RowRelative:=False 
 InsertHyperlink Name:="https://MSDN", Address:="https://msdn.microsoft.com/", SubAddress:="", ScreenTip:="" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]