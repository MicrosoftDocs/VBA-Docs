---
title: TreeviewControl.Hidden property (Excel)
keywords: vbaxl10.chm666073
f1_keywords:
- vbaxl10.chm666073
ms.prod: excel
api_name:
- Excel.TreeviewControl.Hidden
ms.assetid: 134a3b6b-492b-6813-cd40-ce1ff3b52c6c
ms.date: 05/18/2019
localization_priority: Normal
---


# TreeviewControl.Hidden property (Excel)

Returns or sets a **Variant** value that represents the hidden status of the cube field members in the hierarchical member-selection control of a cube field.


## Syntax

_expression_.**Hidden**

_expression_ A variable that represents a **[TreeviewControl](Excel.TreeviewControl.md)** object.


## Remarks

Don't confuse this property with the **[FormulaHidden](Excel.Range.FormulaHidden.md)** property.

The **Hidden** property returns or sets an array. Each element of the array corresponds to a level of the cube field that is hidden. The maximum number of elements is the number of levels in the cube field. Each element of the array is an array of type **String**, which contains unique member names that are hidden at the corresponding level of the control. 

To determine when members are visible (expanded) in the control, see the **[DrilledDown](Excel.PivotItem.DrilledDown.md)** property of the **PivotItem** object. 


## Example

This example hides the second level member [state].[states].[CA].[Covelo] of the first cube field in the first PivotTable report.

```vb
ActiveSheet.PivotTables("PivotTable1").CubeFields(1) _ 
 .TreeviewControl.Hidden = _ 
 Array(Array(""), Array(""), _ 
 Array("[state].[states].[CA].[Covelo]"))
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]