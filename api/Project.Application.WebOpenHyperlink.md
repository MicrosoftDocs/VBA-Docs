---
title: Application.WebOpenHyperlink method (Project)
keywords: vbapj.chm1311
f1_keywords:
- vbapj.chm1311
ms.prod: project-server
api_name:
- Project.Application.WebOpenHyperlink
ms.assetid: f1da5d5f-45a1-02e0-8783-7f919578e3fe
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WebOpenHyperlink method (Project)

Opens the document specified by a hyperlink address. 


## Syntax

_expression_. `WebOpenHyperlink`( `_Address_`, `_SubAddress_`, `_AddHistory_`, `_NewWindow_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Address_|Optional|**String**|The address of the target document. If  **Address** is omitted, the text of the selected field is used.|
| _SubAddress_|Optional|**String**|A location within the target document.|
| _AddHistory_|Optional|**Boolean**|**True** if the target document is added to the History folder. The default value is **True**.|
| _NewWindow_|Optional|**Boolean**|**True** if the target document displays in a new window. The default value is **False**.|

## Return value

 **Boolean**


## Remarks

The  **WebOpenHyperlink** method is only available when the selected assignment, resource, or task field contains a hyperlink.


## Example

The following example inserts a hyperlink in the Gantt Chart and then opens it.


```vb
Sub WebOpen_Hyperlink() 
 
 'Activate Gantt Chart 
 ViewApply Name:="&Gantt Chart" 
 SelectRow Row:=2, RowRelative:=False 
 InsertHyperlink Name:="https://MSDN/", Address:="https://msdn.microsoft.com/", SubAddress:="", ScreenTip:="" 
 
 'Open the webpage 
 WebOpenHyperlink Address:="https://msdn.microsoft.com/", SubAddress:="" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]