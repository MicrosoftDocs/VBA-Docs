---
title: Tables.Add method (Project)
ms.prod: project-server
api_name:
- Project.Tables.Add
ms.assetid: 595c0cb8-fd3f-8f5c-3eaf-588f41dc36dc
ms.date: 06/08/2017
localization_priority: Normal
---


# Tables.Add method (Project)

Adds a  **Table** object to a **Tables** collection.


## Syntax

_expression_.**Add** (_Name_, _Field_, _Task_)

_expression_ A variable that represents a 'Tables' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the table.|
| _Field_|Required|**Long**|The name of the field. Can be one of the  **[PjField](Project.PjField.md)** constants.|
| _Task_|Optional|**Boolean**|**True** if the table being added is a task table. The default value is **True**.|

## Return value

 **Table**


## See also


[Tables Collection Object](Project.tables.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]