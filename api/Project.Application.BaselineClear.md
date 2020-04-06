---
title: Application.BaselineClear method (Project)
keywords: vbapj.chm2384
f1_keywords:
- vbapj.chm2384
ms.prod: project-server
api_name:
- Project.Application.BaselineClear
ms.assetid: a319fc88-2421-eafa-e498-4a0a5f173394
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BaselineClear method (Project)

Clears the baseline data from the baseline fields or clears the data from a  **Start _n_** / **Finish _n_** pair of dates.


## Syntax

_expression_. `BaselineClear`( `_All_`, `_From_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _All_|Optional|**Boolean**|**True** if all tasks in the active project should be cleared. **False** if only the selected tasks should be cleared. The default value is **True**.|
| _From_|Optional|**Long**|The fields to be cleared. The default value is  **pjIntoBaseline**. Can be one of the [PjSaveBaselineTo](Project.PjSaveBaselineTo.md) constants.|

## Return value

 **Boolean**


## Example

The following example first saves the baseline and then clears it.


```vb
Sub Baseline_Clear() 
 
 Dim Result As Boolean 
 
 'Save baseline 
 Result = BaselineSave(True) 
 'Clear baseline 
 Result = BaselineClear (True) 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]