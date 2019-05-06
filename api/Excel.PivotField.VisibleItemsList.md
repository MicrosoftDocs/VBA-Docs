---
title: PivotField.VisibleItemsList property (Excel)
keywords: vbaxl10.chm240146
f1_keywords:
- vbaxl10.chm240146
ms.prod: excel
api_name:
- Excel.PivotField.VisibleItemsList
ms.assetid: ddcc2dce-30bf-ba50-22fa-a4baf41129f5
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotField.VisibleItemsList property (Excel)

Returns or sets a **Variant** specifying an array of strings that represent included items in a manual filter applied to a PivotField. Read/write.


## Syntax

_expression_.**VisibleItemsList**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

This property is applicable to OLAP PivotTables only.


## Example

This example shows manual, inclusive filtering in an OLAP PivotTable.

```vb
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] & _ 
.[Country]").VisibleItemsList = Array("[Customer].[Customer Geography].[Country].&[Australia]") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] & _ 
.[State-Province]").VisibleItemsList = Array("") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] & _ 
.[City]").VisibleItemsList = Array("") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] & _ 
.[Postal Code]").VisibleItemsList = Array("") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] & _ 
.[Full Name]").VisibleItemsList = Array("") 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]