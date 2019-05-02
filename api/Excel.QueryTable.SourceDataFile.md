---
title: QueryTable.SourceDataFile property (Excel)
keywords: vbaxl10.chm518132
f1_keywords:
- vbaxl10.chm518132
ms.prod: excel
api_name:
- Excel.QueryTable.SourceDataFile
ms.assetid: c6fb30b8-c909-7509-65bc-f6df9a3640c6
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.SourceDataFile property (Excel)

Returns or sets a **String** value that indicates the source data file for a query table.


## Syntax

_expression_.**SourceDataFile**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

For file-based data sources (such as Access), the **SourceDataFile** property contains a fully qualified path to the source data file. It is set to **Null** for server-based data sources (such as SQL Server). The **SourceDataFile** property is set to **Null** if the **[Connection](Excel.QueryTable.Connection.md)** property is changed programmatically.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **SourceDataFile** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]