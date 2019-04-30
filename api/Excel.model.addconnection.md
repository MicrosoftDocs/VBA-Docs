---
title: Model.AddConnection method (Excel)
ms.prod: excel
ms.assetid: 58ed2796-9cfa-2737-43c0-f5a5a4badcc3
ms.date: 04/30/2019
localization_priority: Normal
---


# Model.AddConnection method (Excel)

Adds a new **[WorkbookConnection](Excel.WorkbookConnection.md)** to the model with the same properties as the one supplied as an argument.


## Syntax

_expression_.**AddConnection** (_ConnectionToDataSource_)

_expression_ A variable that represents a **[Model](Excel.Model.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ConnectionToDataSource_|Required|WORKBOOKCONNECTION|The Workbook connection.|

## Return value

**WORKBOOKCONNECTION**


## Remarks

This method only works on legacy/non-model external connections and will fail with a run-time error if called with an external model connection as its argument. When calling this method, a new model connection is created and it is named the same as the legacy connection with the existing logic for making the name unique applied (integer at the end).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]