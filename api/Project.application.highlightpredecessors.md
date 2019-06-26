---
title: Application.HighlightPredecessors method (Project)
keywords: vbapj.chm147
f1_keywords:
- vbapj.chm147
ms.prod: project-server
ms.assetid: e4c51516-2e5d-3ef9-3165-84fe6f9ad38b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.HighlightPredecessors method (Project)
Sets or clears task predecessor highlighting for the task path feature.

## Syntax

_expression_. `HighlightPredecessors` _(Set)_

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Set_|Optional|**Variant**|**True** to set task predecessor highlighting; **False** to clear the task predecessor highlighting.|
| _Set_|Optional|**Variant**||
|Name|Required/Optional|Data type|Description|

## Return value

 **Boolean**


## Remarks

The  **HighlightPredecessors** method corresponds to the **Predecessors** item in the **Task Path** drop-down list, on the **FORMAT** tab, under **GANTT CHART TOOLS** on the ribbon.


## Example

Create a project where task 2 is a predecessor of task 3, and then run the following statements in the Immediate window of the VBE. The **PathPredecessor** statement prints **True**.


```vb
Application.SelectRow Row:=2, RowRelative:=False 
Application.HighlightPredecessors True
? ActiveProject.Tasks(3).PathPredecessor
```


## See also


[Application Object](Project.Application.md)



[Task.PathPredecessor Property](Project.task.pathpredecessor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]