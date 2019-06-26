---
title: Application.GanttBarSize method (Project)
keywords: vbapj.chm2058
f1_keywords:
- vbapj.chm2058
ms.prod: project-server
api_name:
- Project.Application.GanttBarSize
ms.assetid: 691ee987-a62b-bf5f-0088-0f153aa64966
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GanttBarSize method (Project)

Sets the height, in [points](../language/glossary/vbe-glossary.md#point), of the Gantt bars in the active Gantt Chart.


## Syntax

_expression_. `GanttBarSize`( `_Size_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Size_|Required|**Long**|A constant specifying the height, in [points](../language/glossary/vbe-glossary.md#point), of the Gantt bars in the active Gantt Chart. Can be one of the following  **[PjBarSize](Project.PjBarSize.md)** constants.|

## Return value

 **Boolean**


## Example

The following example set the bar size to pjBarSize24.


```vb
Sub GanttBar_Size() 
 
 'Activate Gantt Chart view 
 ViewApply Name:="&Gantt Chart" 
 GanttBarSize Size:= 
pjBarSize24
```


```vb
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]