---
title: Application.SetResourceFieldByID method (Project)
keywords: vbapj.chm96
f1_keywords:
- vbapj.chm96
ms.prod: project-server
api_name:
- Project.Application.SetResourceFieldByID
ms.assetid: 1309ee61-6b66-db45-ed69-b0b3dd9b8dda
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SetResourceFieldByID method (Project)

Sets the value of a resource field specified by the field identification number.


## Syntax

_expression_. `SetResourceFieldByID`( `_FieldID_`, `_Value_`, `_AllSelectedResources_`, `_Create_`, `_ResourceID_`, `_ProjectName_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**PjField**|Field identification number. Can be one of the resource fields specified by a **[PjField](Project.PjField.md)** constant or a number returned by the **[FieldNameToFieldConstant](Project.Application.FieldNameToFieldConstant.md)** method.|
| _Value_|Required|**String**|The value of the resource field.|
| _AllSelectedResources_|Optional|**Boolean**|**True** if the value of the field is set for all selected resources. **False** if the value is set for the active resource. The default value is **False**.|
| _Create_|Optional|**Boolean**|**True** if Project should create a resource if the active cell is on an empty row. The default value is **True**.|
| _ResourceID_|Optional|**Long**|The identification number of the resource containing the field to set. If AllSelectedResources is **True**, ResourceID is ignored.|
| _ProjectName_|Optional|**String**|If the active project is a consolidated project, specifies the name of the project for the resource specified by  _ResourceID_. If  _ResourceID_ is not specified, _ProjectName_ is ignored. The default value is the name of the active project.|

## Return value

 **Boolean**


## Remarks

To set a resource field by name, use the **[SetResourceField](Project.Application.SetResourceField.md)** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]