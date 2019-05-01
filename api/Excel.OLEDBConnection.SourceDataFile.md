---
title: OLEDBConnection.SourceDataFile property (Excel)
keywords: vbaxl10.chm794092
f1_keywords:
- vbaxl10.chm794092
ms.prod: excel
api_name:
- Excel.OLEDBConnection.SourceDataFile
ms.assetid: ddadf399-3f93-bd20-22cf-5f9303704218
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBConnection.SourceDataFile property (Excel)

Returns or sets a **String** indicating the source data file for an OLE DB connection. Read/write.


## Syntax

_expression_.**SourceDataFile**

_expression_ A variable that represents an **[OLEDBConnection](Excel.OLEDBConnection.md)** object.


## Remarks

For file-based data sources (for example, Access) the **SourceDataFile** property contains a fully qualified path to the source data file. It is **null** for server-based data sources (for example, SQL Server). The **SourceDataFile** property is set to **null** if the **[Connection](Excel.OLEDBConnection.Connection.md)** property is changed programmatically.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]