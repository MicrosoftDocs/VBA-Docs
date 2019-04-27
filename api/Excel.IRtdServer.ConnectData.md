---
title: IRtdServer.ConnectData method (Excel)
keywords: vbaxl10.chm500006
f1_keywords:
- vbaxl10.chm500006
ms.prod: excel
api_name:
- Excel.IRtdServer.ConnectData
ms.assetid: 2d660ccc-fca7-c794-61f1-4e0578cc7511
ms.date: 04/27/2019
localization_priority: Normal
---


# IRtdServer.ConnectData method (Excel)

Adds new topics from a real-time data (RTD) server. The **ConnectData** method is called when a file is opened that contains real-time data functions or when a user types in a new formula that contains the RTD function.


## Syntax

_expression_.**ConnectData** (_TopicID_, _Strings()_, _GetNewValues_)

_expression_ A variable that represents an **[IRtdServer](Excel.IRtdServer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TopicID_|Required| **Long**| A unique value, assigned by Microsoft Excel, that identifies the topic.|
| _Strings()_|Required| **Variant**|A single-dimensional array of strings identifying the topic.|
| _GetNewValues_|Required| **Boolean**| **True** to determine if new values are to be acquired.|

## Return value

Variant




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]