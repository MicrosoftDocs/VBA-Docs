---
title: Application.HighlightDrivenSuccessors method (Project)
keywords: vbapj.chm150
f1_keywords:
- vbapj.chm150
ms.prod: project-server
ms.assetid: 2c93505b-541f-15a7-31ff-fcddcfa0bb55
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.HighlightDrivenSuccessors method (Project)
Sets or clears task driven successor highlighting for the task path feature.

## Syntax

_expression_. `HighlightDrivenSuccessors` _(Set)_

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Set_|Optional|**Variant**|**True** to set task driven successor highlighting; **False** to clear the task driven successor highlighting.|
| _Set_|Optional|**Variant**||
|Name|Required/Optional|Data type|Description|

## Return value

 **Boolean**


## Remarks

The **HighlightDrivenSuccessors** method corresponds to the **Driven Successors** item in the **Task Path** drop-down list, on the **FORMAT** tab, under **GANTT CHART TOOLS** on the ribbon.


## Example

Create a project where task 4 is a driven successor of task 3, and then run the following statements in the Immediate window of the VBE. The **PathDrivenSuccessor** statement prints **True**.


```vb
Application.SelectRow Row:=3, RowRelative:=False 
Application.HighlightDrivenSuccessors True
? ActiveProject.Tasks(4).PathDrivenSuccessor
```


## See also


[Application Object](Project.Application.md)



[Task.PathDrivenSuccessor Property](Project.task.pathdrivensuccessor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]