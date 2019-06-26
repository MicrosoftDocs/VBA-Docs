---
title: Application.SetRowHeight method (Project)
keywords: vbapj.chm2118
f1_keywords:
- vbapj.chm2118
ms.prod: project-server
api_name:
- Project.Application.SetRowHeight
ms.assetid: bfa4a87b-9e9f-9937-4b9d-a7b26576a5da
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SetRowHeight method (Project)

Sets the height of the specified rows.


## Syntax

_expression_. `SetRowHeight`( `_Unit_`, `_Rows_`, `_UseUniqueID_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional|**Integer**|The height of the rows, in lines. The maximum value for Unit is 20.|
| _Rows_|Optional|**String**|The row(s) to select. The value for Rows can be a single row (for example, "5"), a range of rows (for example, "1-8"), a list of discontiguous rows (for example, "5,7-9,12"), or "ALL" to select every row. If Rows is not specified and an existing selection exists, the selection will be used. The default with no existing selection is to use the active row.|
| _UseUniqueID_|Optional|**Boolean**|**True** if the value specified with Rows is the unique identification number(s) for resources or tasks. **False** if Rows specifies row numbers. The default value is **False**.|

## Return value

 **Boolean**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]