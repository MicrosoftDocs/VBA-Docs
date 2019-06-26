---
title: Application.ProjectBeforeTaskChange2 event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeTaskChange2
ms.assetid: 00992e39-dcbd-3826-4ce6-e2be55dc9c2c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeTaskChange2 event (Project)

Occurs before the user changes the value of a task field. Uses the **EventInfo** object parameter.


## Syntax

_expression_.**ProjectBeforeTaskChange2** (_tsk_, _Field_, _NewVal_, _Info_)

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _tsk_|Required|**Task**|The task whose field is being changed.|
| _Field_|Required|**PjField**|The field being changed. If more than one field is changed by the user, the event is fired for each field changed. Can be one of the **[PjField](project.pjfield.md)** constants.|
| _NewVal_|Required|**Variant**|The new value for the field specified with _Field_.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is **False** when the event occurs. If the event procedure sets this argument to **True**, the value for the field specified with _Field_ is not changed.|

## Return value

Nothing


## Remarks

Project events do not occur when the project is embedded in another document or application. For more information and sample code for creating and testing an event handler, see [Using Events with Application and Project Objects](../project/Concepts/using-events-with-application-and-project-objects.md).

The **ProjectBeforeTaskChange2** event doesn't occur when timescaled data changes, when constraint data in the Task Details Form changes, when a task is split by manipulating its task bar on the Gantt Chart, when changes are made to outline level or outline number, when a baseline is saved, when a baseline is cleared, when an entire task row is pasted, during resource pool operations, when inserting or removing a subproject, or when changes have been made by using a custom form.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]