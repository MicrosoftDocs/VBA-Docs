---
title: QueryTable.AfterRefresh event (Excel)
keywords: vbaxl10.chm519074
f1_keywords:
- vbaxl10.chm519074
ms.prod: excel
api_name:
- Excel.QueryTable.AfterRefresh
ms.assetid: 91d930e3-4360-4ec2-8772-dcd67c9e8c41
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.AfterRefresh event (Excel)

Occurs after a query is completed or canceled.


## Syntax

_expression_.**AfterRefresh** (_Success_)

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Success_|Required| **Boolean**| **True** if the query was completed successfully.|

## Return value

Nothing


## Example

This example uses the _Success_ argument to determine which section of code to run.

```vb
Private Sub QueryTable_AfterRefresh(Success As Boolean) 
 If Success Then 
 ' Query completed successfully 
 Else 
 ' Query failed or was cancelled 
 End If 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
