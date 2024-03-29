---
title: PivotField.PropertyOrder property (Excel)
keywords: vbaxl10.chm240133
f1_keywords:
- vbaxl10.chm240133
api_name:
- Excel.PivotField.PropertyOrder
ms.assetid: b938d2bd-3e64-a861-c058-96daa81830bf
ms.date: 05/07/2019
ms.localizationpriority: medium
---


# PivotField.PropertyOrder property (Excel)

Valid only for PivotTable fields that are member property fields. Returns a **Long** indicating the display position of the member property within the cube field to which it belongs. Read/write.


## Syntax

_expression_.**PropertyOrder**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

Setting this property will rearrange the order of the properties for this cube field. This property is one-based. The allowable range is from one to the maximum number of member property fields being displayed for the hierarchy. 

If the **[IsMemberProperty](Excel.PivotField.IsMemberProperty.md)** property is **False**, using the **PropertyOrder** property creates a run-time error.


## Example

This example determines if there are member properties in the fourth field, and if there are, displays the position of the member properties. Depending on the findings, Excel notifies the user. This example assumes that a PivotTable exists on the active worksheet and that it is based on an Online Analytical Processing (OLAP) data source.

```vb
Sub CheckPropertyOrder() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields(4) 
 
 ' Check for member properties and notify user. 
 If pvtField.IsMemberProperty = False Then 
 MsgBox "No member properties present." 
 Else 
 MsgBox "The property order of the members is: " & _ 
 pvtField.PropertyOrder 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]