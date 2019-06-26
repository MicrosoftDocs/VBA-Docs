---
title: Application.ProjectBeforeResourceChange2 event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeResourceChange2
ms.assetid: 84128c94-0d0d-f8f2-6d5a-4c05a61a0a8d
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectBeforeResourceChange2 event (Project)

Occurs before the user changes the value of a resource field. Uses the **EventInfo** object parameter.


## Syntax

_expression_.**ProjectBeforeResourceChange2** (_res_, _Field_, _NewVal_, _Info_)

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _res_|Required|**Resource**|The resource whose field is being changed.|
| _Field_|Required|**Long**|The field being changed. If more than one field is changed by the user, the event is fired for each field changed. Can be one of the **[PjField](project.pjfield.md)** constants.|
| _NewVal_|Required|**Variant**|The new value for the field specified with Field.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is **False** when the event occurs. If the event procedure sets this argument to **True**, the value for the field specified with Field is not changed.|



## Return value

Nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.

The **ProjectBeforeResourceChange2** event doesn't occur when timescaled data changes, when a baseline is cleared, when an entire resource row is pasted, during resource pool operations, when inserting or removing a subproject, or when changes have been made by using a custom form.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]