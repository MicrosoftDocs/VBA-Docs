---
title: RecentFiles.Add method (Excel)
keywords: vbaxl10.chm172077
f1_keywords:
- vbaxl10.chm172077
ms.prod: excel
api_name:
- Excel.RecentFiles.Add
ms.assetid: 70d4c4e0-b0f5-8143-0f23-69dc1c85736e
ms.date: 05/11/2019
localization_priority: Normal
---


# RecentFiles.Add method (Excel)

Adds a file to the list of recently used files.


## Syntax

_expression_.**Add** (_Name_)

_expression_ A variable that represents a **[RecentFiles](Excel.RecentFiles.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The file name.|

## Return value

A **[RecentFile](Excel.RecentFile.md)** object contained by the collection.


## Example

This example adds Oscar.xls to the list of recently used files.

```vb
Application.RecentFiles.Add Name:="Oscar.xls"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]