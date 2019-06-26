---
title: LookupTable.AddChild method (Project)
keywords: vbapj.chm132387
f1_keywords:
- vbapj.chm132387
ms.prod: project-server
api_name:
- Project.LookupTable.AddChild
ms.assetid: 6e7d3a9c-8a71-26f8-628a-2efff5897951
ms.date: 06/08/2017
localization_priority: Normal
---


# LookupTable.AddChild method (Project)

Adds a lookup table entry to a **[LookupTable](Project.lookuptable.md)** collection. Returns a reference to the **[LookupTableEntry](Project.LookupTableEntry.md)**.


## Syntax

_expression_.**AddChild** (_Name_, _ParentUniqueID_)

_expression_ A variable that represents a 'LookupTable' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the lookup table entry.|
| _ParentUniqueID_|Optional|**Long**|If this value is not specified, the entry is inserted at the top level. Otherwise, the entry is inserted as the child of the entry with the specified unique identifier (UID). The method ensures that the entry with the specified UID is in the correct lookup table.|

## Return value

 **LookupTableEntry**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]