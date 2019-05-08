---
title: PivotTable.AddDataField method (Excel)
keywords: vbaxl10.chm235142
f1_keywords:
- vbaxl10.chm235142
ms.prod: excel
api_name:
- Excel.PivotTable.AddDataField
ms.assetid: 768b1eb7-80ea-fb0f-0de5-803ec19bbe18
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.AddDataField method (Excel)

Adds a data field to a PivotTable report. Returns a **[PivotField](Excel.PivotField.md)** object that represents the new data field.


## Syntax

_expression_.**AddDataField** (_Field_, _Caption_, _Function_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Field_|Required| **Object**|The unique field on the server. If the source data is Online Analytical Processing (OLAP), the unique field is a cube field. If the source data is non-OLAP (non-OLAP source data), the unique field is a PivotTable field.|
| _Caption_|Optional| **Variant**|The label used in the PivotTable report to identify this data field.|
| _Function_|Optional| **Variant**|The function performed in the added data field.|

## Return value

PivotField


## Example

This example adds a data field titled Total Score to a PivotTable called PivotTable1. This example assumes that a table exists in which one of the columns contains a column titled Score.

```vb
Sub AddMoreFields() 
 
 With ActiveSheet.PivotTables("PivotTable1") 
 .AddDataField ActiveSheet.PivotTables( _ 
 "PivotTable1").PivotFields("Score"), "Total Score" 
 End With 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
