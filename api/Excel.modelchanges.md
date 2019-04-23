---
title: ModelChanges object (Excel)
keywords: vbaxl10.chm959072
f1_keywords:
- vbaxl10.chm959072
ms.prod: excel
ms.assetid: fd2388eb-48ab-c238-2ffa-8c3f6d20fe36
ms.date: 03/30/2019
localization_priority: Normal
---


# ModelChanges object (Excel)

Represents changes made to the data model. 


## Remarks

The **ModelChanges** object contains information about which changes were made to the data model when the **[ModelChange](Excel.workbook.modelchange.md)** event of the **Workbook** object occurs after a model operation. 

When Micrososft Excel makes changes to the data model, multiple changes can be made in the same operation, and the **ModelChanges** object will include information about all the changes made in one model operation.

## Properties

- [Application](Excel.modelchanges.application.md)
- [ColumnsAdded](Excel.modelchanges.columnsadded.md)
- [ColumnsChanged](Excel.modelchanges.columnschanged.md)
- [ColumnsDeleted](Excel.modelchanges.columnsdeleted.md)
- [Creator](Excel.modelchanges.creator.md)
- [MeasuresAdded](Excel.modelchanges.measuresadded.md)
- [Parent](Excel.modelchanges.parent.md)
- [RelationshipChange](Excel.modelchanges.relationshipchange.md)
- [Source](Excel.modelchanges.source.md)
- [TableNamesChanged](Excel.modelchanges.tablenameschanged.md)
- [TablesAdded](Excel.modelchanges.tablesadded.md)
- [TablesDeleted](Excel.modelchanges.tablesdeleted.md)
- [TablesModified](Excel.modelchanges.tablesmodified.md)
- [UnknownChange](Excel.modelchanges.unknownchange.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
