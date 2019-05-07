---
title: Shape.CellsSRCExists property (Visio)
keywords: vis_sdr.chm11213210
f1_keywords:
- vis_sdr.chm11213210
ms.prod: visio
api_name:
- Visio.Shape.CellsSRCExists
ms.assetid: 7d614820-2a64-c3ee-b61c-a7c0dcfb90c8
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.CellsSRCExists property (Visio)

Determines whether a ShapeSheet cell exists in the scope of a search. Read-only.


## Syntax

_expression_. `CellsSRCExists`( `_Section_` , `_Row_` , `_Column_` , `_fExistsLocally_` )

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The cell's section index.|
| _Row_|Required| **Integer**|The cell's row index.|
| _Column_|Required| **Integer**|The cell's column index.|
| _fExistsLocally_|Required| **Integer**|The scope of the search.|

## Return value

Integer


## Remarks

Constants for section, row, and column indices are declared by the Visio type library as members of  **[VisSectionIndices](Visio.vissectionindices.md)**, **[VisRowIndices](Visio.visrowindices.md)**, and **[VisCellIndices](Visio.viscellindices.md)**, respectively.

The  _fExistsLocally_ argument specifies the scope of the search:




- If  _fExistsLocally_ is non-zero (**True**), the **CellsSRCExists** property returns **True** only if the object contains the cell locally; if the cell is inherited, the **CellsSRCExists** property returns **False**.
    
- If  _fExistsLocally_ is zero (**False**), the **CellsSRCExists** property returns **True** if the object either contains or inherits the cell.
    


To search for a cell by name, use the  **CellExists** or **CellExistsU** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]