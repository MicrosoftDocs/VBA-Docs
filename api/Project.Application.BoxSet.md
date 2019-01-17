---
title: Application.BoxSet method (Project)
keywords: vbapj.chm49
f1_keywords:
- vbapj.chm49
ms.prod: project-server
api_name:
- Project.Application.BoxSet
ms.assetid: 06bcae73-5208-824d-4f55-119f35b37718
ms.date: 11/09/2018
localization_priority: Normal
---


# Application.BoxSet method (Project)

Creates, selects, or moves a task in the Network Diagram view.

## Syntax

_expression_. BoxSet( _action_, _TaskID_, _XPosition_, _YPosition_, _ProjectName_ )

_expression_ A variable that represents an [Application](Project.Application.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _action_|Optional|**Long**|The operation to perform on the specified task(s). The default value is **pjBoxSelect**. Can be one of the **[PjBoxSet](Project.PjBoxSet.md)** constants.|
| _TaskID_|Optional|**Long**|The identification number of the task. If **action** is **pjBoxCreate**, **TaskID** is ignored.|
| _XPosition_|Optional|**Long**|The horizontal position of the task, in pixels. Required if **action** is **pjBoxMoveAbsolute** or **pjBoxMoveRelative**.<br/><br/>If **action** is **pjBoxCreate** or **pjBoxMoveAbsolute**, **XPosition** is the absolute horizontal position of the upper-left corner of the task.<br/><br/>If **action** is **pjBoxMoveRelative**, **XPosition** is the amount to move the task horizontally relative to the current position.<br/><br/>If **action** is **pjBoxAddToSelection**, **pjBoxSelect**, or **pjBoxUnselect**, **XPosition** is ignored.|
| _YPosition_|Optional|**Long**|The vertical position of the task, in pixels. Required if **action** is **pjBoxMoveAbsolute** or **pjBoxMoveRelative**.<br/><br/>If **action** is **pjBoxCreate** or **pjBoxMoveAbsolute**, **YPosition** is the absolute vertical position of the upper-left corner of the task.<br/><br/>If **action** is **pjBoxMoveRelative**, **YPosition** is the amount to move the task vertically relative to the current position.<br/><br/>If **action** is **pjBoxAddToSelection**, **pjBoxSelect**, or **pjBoxUnselect**, **YPosition** is ignored.|
| _ProjectName_|Optional|**String**|If the active project is a consolidated project, specifies the name of the project for the identification number specified by **TaskID**. If **TaskID** is not specified, **ProjectName** is ignored. The default value is the name of the active project.|

## Return value

**Boolean**

## Remarks

If only one task box is selected, specifying **pjBoxUnselect** has no effect.

If automatic layout has been activated for the Network Diagram view, **XPosition** and **YPosition** have no effect.

## Example

The following example adds the task with TaskID 2 to the selected tasks.

```vb
Sub Box_Set() 
 
 'Activate the Network Diagram view 
 ViewApply Name:="Network Diagram" 
 BoxSet action:=pjBoxAddToSelection, TaskID:="2" 
End Sub
```

> [!NOTE] 
> BoxSet does not currently work for subprojects. You can place the subproject name in the Project Name attribute and set the TaskID, but it does not perform the action to the box from the subproject in the Network Diagram. 
