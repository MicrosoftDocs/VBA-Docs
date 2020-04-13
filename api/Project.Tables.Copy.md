---
title: Tables.Copy method (Project)
keywords: vbapj.chm132701
f1_keywords:
- vbapj.chm132701
ms.prod: project-server
api_name:
- Project.Tables.Copy
ms.assetid: dfc2f25b-e60c-ef25-9e7c-2808ce76a4ba
ms.date: 06/08/2017
localization_priority: Normal
---


# Tables.Copy method (Project)

Makes a copy of a group definition for the **Tables** collection and returns a reference to the **[Table](Project.Table.md)** object.


## Syntax

_expression_.**Copy** (_Source_, _NewName_)

_expression_ A variable that represents a 'Tables' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Required|**String**|The name of the table to copy.|
| _NewName_|Required|**String**|The name of the new table.|

## Return value

 **Table**


## See also


[Tables Collection Object](Project.tables.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]