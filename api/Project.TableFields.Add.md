---
title: TableFields.Add method (Project)
keywords: vbapj.chm132691
f1_keywords:
- vbapj.chm132691
ms.prod: project-server
api_name:
- Project.TableFields.Add
ms.assetid: d4e6af9f-6d95-49f0-8828-dcd39dbb9f13
ms.date: 06/08/2017
localization_priority: Normal
---


# TableFields.Add method (Project)

Adds a  **TableField** object to a **TableFields** collection.


## Syntax

_expression_.**Add** (_Field_, _AlignData_, _Width_, _Title_, _AlignTitle_, _Before_, _AutoWrap_)

_expression_ A variable that represents a 'TableFields' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Field_|Required|**Long**|The name of the table field. Can be one of the  **[PjField](Project.PjField.md)** constants.|
| _AlignData_|Optional|**Long**|The alignment of the table data. Can be one of the  **[PjAlignment](Project.PjAlignment.md)** constants. The default value is **pjRight**.|
| _Width_|Optional|**Long**|The width of the table field in points. The default value is 10.|
| _Title_|Optional|**String**|The title of the table field.|
| _AlignTitle_|Optional|**Long**|The alignment of the title. Can be one of the  **PjAlignment** constants. The default value is **pjCenter**.|
| _Before_|Optional|**Long**|Position of the title. The default value is -1.|
| _AutoWrap_|Optional|**Boolean**|**True** if the data in the table field automatically wrap. The default value is **True**.|

## Return value

 **TableField**


## See also


[TableFields Collection Object](Project.tablefields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]