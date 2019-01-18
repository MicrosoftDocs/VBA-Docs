---
title: PivotField.DataType property (Excel)
keywords: vbaxl10.chm240079
f1_keywords:
- vbaxl10.chm240079
ms.prod: excel
api_name:
- Excel.PivotField.DataType
ms.assetid: 95671f37-9886-822f-672c-1c5706b9c0bf
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotField.DataType property (Excel)

Returns a  **[xlPivotFieldDataType](Excel.XlPivotFieldDataType.md)** value that represents the type of data in the PivotTable field.


## Syntax

_expression_. `DataType`

_expression_ A variable that represents a [PivotField](Excel.PivotField.md) object.


## Example

This example displays the data type of the field named "ORDER_DATE."


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
Select Case pvtTable.PivotFields("ORDER_DATE").DataType 
 Case Is = xlText 
 MsgBox "The field contains text data" 
 Case Is = xlNumber 
 MsgBox "The field contains numeric data" 
 Case Is = xlDate 
 MsgBox "The field contains date data" 
End Select
```


## See also


[PivotField Object](Excel.PivotField.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]