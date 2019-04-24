---
title: QueryTable.SourceConnectionFile property (Excel)
keywords: vbaxl10.chm518131
f1_keywords:
- vbaxl10.chm518131
ms.prod: excel
api_name:
- Excel.QueryTable.SourceConnectionFile
ms.assetid: 2f7472a2-dbac-5dbb-ea27-1508211f001f
ms.date: 06/08/2017
localization_priority: Normal
---


# QueryTable.SourceConnectionFile property (Excel)

Returns or sets a  **String** indicating the Microsoft Office Data Connection file or similar file that was used to create the QueryTable. Read/write.


## Syntax

_expression_. `SourceConnectionFile`

_expression_ A variable that represents a [QueryTable](Excel.QueryTable.md) object.


## Remarks

Data from Web queries or text queries is imported as a  **[QueryTable](Excel.QueryTable.md)** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object. You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **SourceConnectionFile** property.

If you import data using the user interface, data from a web query or a text query is imported as a  **[QueryTable](Excel.QueryTable.md)** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data using the object model, data from a web query or a text query must be imported as a  **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the  **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **SourceConnectionFile** property.


## See also


[QueryTable Object](Excel.QueryTable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]