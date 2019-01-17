---
title: xlCellChangedState enumeration (Excel)
ms.prod: excel
api_name:
- Excel.xlCellChangedState
ms.assetid: d0242314-afe9-f5e0-6c54-65ca7b4fb800
ms.date: 06/08/2017
localization_priority: Normal
---


# xlCellChangedState enumeration (Excel)

Specifies whether a PivotTable value cell has been edited or recalculated since the PivotTable report was created or the last commit operation was performed. 



|Name|Value|Description|
|:-----|:-----|:-----|
| **xlCellChangeApplied**|3|The value in the cell has been edited or recalculated, and that change has been applied to the data source. (Applies only PivotTable reports with OLAP data sources)|
| **xlCellChanged**|2|The value in the cell has been edited or recalculated.|
| **xlCellNotChanged**|1|The value in the cell has not been edited or recalculated.|

## Remarks

Applying and saving changes applies only to PivotTable reports with OLAP data sources. For more information about the meaning of the  **xlCellChangedState** enumeration constant values, see the **[CellChanged](Excel.PivotCell.CellChanged.md)** property of the **[PivotCell](Excel.PivotCell.md)** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]