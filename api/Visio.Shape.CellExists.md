---
title: Shape.CellExists Property (Visio)
keywords: vis_sdr.chm11213185
f1_keywords:
- vis_sdr.chm11213185
ms.prod: visio
api_name:
- Visio.Shape.CellExists
ms.assetid: 479c4d99-0282-3ab0-2e6f-4a17e48adfab
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.CellExists Property (Visio)

Determines whether a particular ShapeSheet cell exists in the scope of the search. Read-only.


## Syntax

 _expression_. `CellExists`( `_localeSpecificCellName_` , `_fExistsLocally_` )

 _expression_ A variable that represents a [Shape](./Visio.Shape.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _localeSpecificCellName_|Required| **String**|The local or universal name of the ShapeSheet cell for which you want to search.|
| _fExistsLocally_|Required| **Integer**|The scope of the search.|

## Return value

Integer


## Remarks

The  _localeSpecificCellName_ argument can specify a cell name in either local or universal syntax. To search for a cell by section, row, and column index, use the **CellsSRCExists** property.

The  _fExistsLocally_ argument specifies the scope of the search.




- If  _fExistsLocally_ is non-zero (**True**), the **CellExists** property value is **True** only if the object contains the cell locally; if the cell is inherited, the **CellExists** property value is **False**.
    
- If  _fExistsLocally_ is zero (**False**), the **CellExists** property value is **True** if the object either contains or inherits the cell.
    


For a list of cell index values, view the Visio type library for the members of class  **[VisCellIndices](Visio.viscellindices.md)**.




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **CellExists** property to determine if a cell exists by passing either the cell's local name or its universal name. Use the **CellExistsU** property to determine if a cell exists by passing the cell's universal name.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]