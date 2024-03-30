---
title: WorkbookQuery.Delete method (Excel)
keywords: vbaxl10.chm974077
f1_keywords:
- vbaxl10.chm974077
ms.assetid: 05f42f34-1814-870f-081a-c0538b438aec
ms.date: 12/29/2021
ms.localizationpriority: medium
---


# WorkbookQuery.Delete method (Excel)

Deletes this query and its underlying connection and removes it from the **Queries** collection.


## Syntax

_expression_.**Delete**(_DeleteConnection_)

_expression_ A variable that represents a **[WorkbookQuery](Excel.WorkbookQuery.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DeleteConnection_|Optional| **Variant**| **True** To delete the both the query and its underlying connection . The default is **False**.|


## Return value

**Nothing**


## Remarks

By default, the underlying connection is not deleted. To delete both the query and the underlying connection, add the parameter (TRUE).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]