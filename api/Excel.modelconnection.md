---
title: ModelConnection object (Excel)
keywords: vbaxl10.chm921072
f1_keywords:
- vbaxl10.chm921072
ms.prod: excel
ms.assetid: db1b8e2b-76f7-5a6f-b510-6a4d6c4e9857
ms.date: 03/30/2019
localization_priority: Normal
---


# ModelConnection object (Excel)

Contains information for the new Model Connection Type introduced in Excel 2013 to interact with the integrated data model.


## Remarks

Use for all PivotTables connected to the model. The **ModelConnection** object is a new "special" workbook connection that always exists (and cannot be deleted) in workbooks that have a data model. It is created when the model is first created in a workbook. All the properties of this workbook connection are read-only. Its name is _Workbook Data Model_.

## Properties

- [ADOConnection](Excel.modelconnection.adoconnection.md)
- [Application](Excel.modelconnection.application.md)
- [CalculatedMembers](Excel.modelconnection.calculatedmembers.md)
- [CommandText](Excel.modelconnection.commandtext.md)
- [CommandType](Excel.modelconnection.commandtype.md)
- [Creator](Excel.modelconnection.creator.md)
- [Parent](Excel.modelconnection.parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
