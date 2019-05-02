---
title: XlRobustConnect enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlRobustConnect
ms.assetid: 124b8c0f-5120-043e-f226-80d0a7fefe15
ms.date: 05/03/2019
localization_priority: Normal
---


# XlRobustConnect enumeration (Excel)

Specifies how the PivotTable cache or a [query table](excel.querytable.md) connects to its data source.

<br/>

|Name|Value|Description|
|:-----|:-----|:-----|
|**xlAlways** |1|The PivotTable cache or query table always uses external source information (as defined by the **[SourceConnectionFile](Excel.PivotCache.SourceConnectionFile.md)** or **[SourceDataFile](Excel.PivotCache.SourceDataFile.md)** property) to reconnect.|
|**xlAsRequired** |0|The PivotTable cache or query table uses external source information to reconnect by using the **[Connection](Excel.PivotCache.Connection.md)** property.|
|**xlNever** |2|The PivotTable cache or query table never uses source information to reconnect.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]



