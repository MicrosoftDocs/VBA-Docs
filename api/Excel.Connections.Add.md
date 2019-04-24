---
title: Connections.Add method (Excel)
keywords: vbaxl10.chm776079
f1_keywords:
- vbaxl10.chm776079
ms.prod: excel
api_name:
- Excel.Connections.Add
ms.assetid: 2dff072d-b250-e052-64d7-f75a4746a23f
ms.date: 04/23/2019
localization_priority: Normal
---


# Connections.Add method (Excel)

Adds a new connection to the workbook.


## Syntax

_expression_.**Add** (_Name_, _Description_, _ConnectionString_, _CommandText_, _lCmdtype_, _CreateModelConnection_, _ImportRelationships_)

_expression_ A variable that represents a **[Connections](Excel.Connections.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|Name of the connection.|
| _Description_|Required| **String**|Brief description about the connection.|
| _ConnectionString_|Required| **Variant**|The connection string.|
| _CommandText_|Required| **Variant**|The command text to create the connection.|
| _lCmdtype_|Optional| **Variant**|Command type.|
| _CreateModelConnection_|Optional| **Boolean**|Specifies whether to create a connection to the PowerPivot model.|
| _ImportRelationships_|Optional| **Boolean**|Specifies whether to import any existing relationships.|

## Return value

WorkbookConnection




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
