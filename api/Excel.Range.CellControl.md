---
title: CellControl object (Excel)
keywords: 
api_name:
- Excel.CellControl
ms.assetid: 8cf98be1-cd6a-4008-9dd7-bd445ef33977
ms.date: 07/24/2024
ms.localizationpriority: medium
---


# CellControl object (Excel)

This is a child of the Range object, representing the cell controls contained within that range. Properties and methods here let you inspect and modify the cell controls on the range.

## Syntax

Property: Type (read only) 

An enumeration value, used to identify the type of cell control. 

If all cells in the range have the same control type, this returns that type that's shared across all the given cells. Otherwise the cells contain a mix of cell control types, and this returns an empty value.

For example, if A1 has a checkbox and A2 has no cell control, A1:A2 contains a mix of cell control types (one xlTypeCheckbox and one xlTypeNone), and so you'd get an empty result by invoking Range("A1:A2").CellControl.Type. 


## Return values

| Enumeration Value   | Type|
|---------------------|-----------|
|xlTypeNone = 0       |Default type of CellControl. The CellControl object on a range with no cell control formatting will return xlTypeNone as its type.       |
|xlTypeUnknown = 1    |Type associated with a future CellControl object for backward compatibility.<br> For example, in an older Excel version which does not understand Dropdown control (with type xlTypeDropdown), but understands CellControl, will return the CellControl type as xlTypeUnknown of a range with a Dropdown formatting.    |
|xlTypeCheckbox = 2   |Type returned for the Checkbox formatting.   |

## Example

```vb
Range("B1").FormulaR1C1 = Range("A1").CellControl.Type
Range("C1").FormulaR1C1 = Range("A1:A100").CellControl.Type 
```

This example cets the Checkbox formatting on a range. It will overwrite existing cell control formatting if invoked on a range where that's present.

```vb
Range("A1").CellControl.SetCheckbox
Range("A1:A100").CellControl.SetCheckbox 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
