---
title: Application.TableApply method (Project)
keywords: vbapj.chm402
f1_keywords:
- vbapj.chm402
ms.prod: project-server
api_name:
- Project.Application.TableApply
ms.assetid: 3d335475-a0b7-dd61-1c93-a668a878d347
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.TableApply method (Project)

Applies a table to the active view.


## Syntax

_expression_. `TableApply`( `_Name_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**| The name of the table to apply.|

## Return value

 **Boolean**


## Example

The following example applies the Variance table to the active view.


```vb
Sub ApplyVarianceTable() 
 TableApply "Variance" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]