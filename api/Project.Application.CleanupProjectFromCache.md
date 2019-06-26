---
title: Application.CleanupProjectFromCache method (Project)
keywords: vbapj.chm2191
f1_keywords:
- vbapj.chm2191
ms.prod: project-server
api_name:
- Project.Application.CleanupProjectFromCache
ms.assetid: 40fef64a-036f-8e1c-ce86-0c3609777f77
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CleanupProjectFromCache method (Project)

Deletes the specified project file from the local cache. Available only in Project Professional.


## Syntax

_expression_. `CleanupProjectFromCache`( `_FileName_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Optional|**String**|Name of the project file to delete from the cache.|

## Return value

Boolean


## Remarks

You can use the  **CleanupProjectFromCache** method if you suspect the project in the local cache is corrupted. If _FileName_ is omitted, **CleanupProjectFromCache** does nothing.


## Example




```vb
CleanupProjectFromCache("Sample.mpp")
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]