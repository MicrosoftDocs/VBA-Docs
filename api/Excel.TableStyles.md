---
title: TableStyles object (Excel)
keywords: vbaxl10.chm840072
f1_keywords:
- vbaxl10.chm840072
ms.prod: excel
api_name:
- Excel.TableStyles
ms.assetid: 952da370-51cb-b1e0-a413-15cb558099b5
ms.date: 04/02/2019
localization_priority: Normal
---


# TableStyles object (Excel)

Represents styles that can be applied to a table.


## Remarks

Table styles provide a way to format an entire table or PivotTable. Table styles replace the existing auto format feature for formatting an entire table.

Table styles differ from auto format in the following ways:

- You can create and reuse custom table styles.
    
- Table styles work with themes.
    
- Changing the document theme color scheme and/or font scheme will change the look of the built-in table styles.
    
- Table styles can reapply styles to objects such as PivotTables and tables as the object changes. The tables will remember the style applied to them and will re-display appropriately when cells are added, removed, hidden, and shown.
    
- Table styles have a visual user interface in the ribbon.
    

## Methods

- [Add](Excel.TableStyles.Add.md)
- [Item](Excel.TableStyles.Item.md)

## Properties

- [Application](Excel.TableStyles.Application.md)
- [Count](Excel.TableStyles.Count.md)
- [Creator](Excel.TableStyles.Creator.md)
- [Parent](Excel.TableStyles.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]