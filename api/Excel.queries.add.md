---
title: Queries.Add method (Excel)
keywords: vbaxl10.chm976074
f1_keywords:
- vbaxl10.chm976074
ms.assetid: 184711c0-2ce4-ba6e-df56-1f7fdd60ab2c
ms.date: 05/09/2019
ms.prod: excel
localization_priority: Normal
---


# Queries.Add method (Excel)

Adds a new **[WorkbookQuery](Excel.workbookquery.md)** object to the **Queries** collection.


## Syntax

_expression_.**Add** (_Name_, _Formula_, _Description_)

_expression_ A variable that represents a **[Queries](excel.queries.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the query.|
| _Formula_|Required|**String**|The Power Query M formula for the new query.|
| _Description_|Optional|**Variant**|The description of the query.|

## Return value

**WorkbookQuery**


## Example

The following example shows how to add a query to a workbook from an existing CSV file.

```vb
Dim myConnection As WorkbookConnection
Dim mFormula As String
mFormula = _
"let Source = Csv.Document(File.Contents(""C:\data.txt""),null,""#(tab)"",null,1252) in Source"
query1 = ActiveWorkbook.Queries.Add("query1", mFormula)

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
