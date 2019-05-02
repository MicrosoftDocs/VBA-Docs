---
title: QueryTable.BeforeRefresh event (Excel)
keywords: vbaxl10.chm519073
f1_keywords:
- vbaxl10.chm519073
ms.prod: excel
api_name:
- Excel.QueryTable.BeforeRefresh
ms.assetid: 763cfe16-d48c-07f2-73e1-5c59021b4e58
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.BeforeRefresh event (Excel)

Occurs before any refreshes of the query table. This includes refreshes resulting from calling the **Refresh** method, from the user's actions in the product, and from opening the workbook containing the query table.


## Syntax

_expression_.**BeforeRefresh** (_Cancel_)

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the refresh doesn't occur when the procedure is finished.|

## Return value

Nothing


## Example

This example runs before the query table is refreshed.

```vb
Private Sub QueryTable_BeforeRefresh(Cancel As Boolean) 
 a = MsgBox("Refresh Now?", vbYesNoCancel) 
 If a = vbNo Then Cancel = True 
 MsgBox Cancel 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]