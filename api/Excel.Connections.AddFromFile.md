---
title: Connections.AddFromFile method (Excel)
keywords: vbaxl10.chm776080
f1_keywords:
- vbaxl10.chm776080
ms.prod: excel
api_name:
- Excel.Connections.AddFromFile
ms.assetid: 24b13d14-6e5e-ee76-d4a9-79441647a803
ms.date: 04/23/2019
localization_priority: Normal
---


# Connections.AddFromFile method (Excel)

Adds a connection from the specified file.


## Syntax

_expression_.**AddFromFile** (_FileName_, _CreateModelConnection_, _ImportRelationships_)

_expression_ A variable that represents a **[Connections](Excel.Connections.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|Name of the file.|
| _CreateModelConnection_|Optional| **Boolean**|Specifies whether to create the connection to the model.|
| _ImportRelationships_|Optional| **Boolean**|Specifies whether to import the connection relationship.|

## Return value

WorkbookConnection




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]