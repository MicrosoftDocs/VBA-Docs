---
title: Application.OrganizerMoveItem method (Project)
keywords: vbapj.chm127
f1_keywords:
- vbapj.chm127
ms.prod: project-server
api_name:
- Project.Application.OrganizerMoveItem
ms.assetid: a597c657-130e-2e7b-3837-7e3f95421af7
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.OrganizerMoveItem method (Project)

Moves an item in the Organizer.


## Syntax

_expression_. `OrganizerMoveItem`( `_Type_`, `_FileName_`, `_ToFileName_`, `_Name_`, `_Task_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Long**|The type of item to move. Can be one of the **[PjOrganizer](Project.PjOrganizer.md)** constants. The default value is **pjViews**.|
| _FileName_|Required|**String**|The name of the file containing the item to move.|
| _ToFileName_|Required|**String**|The name of the file where the item should be placed.|
| _Name_|Optional|**String**|The name of the item to move. The default is to move all items specified with  **Type**.|
| _Task_|Optional|**Boolean**|**True** if the item applies to tasks. **False** if the item applies to resources. The default value is **True**.|

## Return value

 **Boolean**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]