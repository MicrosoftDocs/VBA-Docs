---
title: TextRange2.InsertChartField method (Office)
ms.assetid: 3ced5d2c-b3a4-6bf3-3d3c-b1145e7b9eab
ms.date: 01/25/2019
ms.prod: office
localization_priority: Normal
---


# TextRange2.InsertChartField method (Office)

Inserts a field into the body of a data label in a chart. 

This method applies only to data labels in a chart. Calling this method on any other kind of **TextRange2** object will raise a run-time error.

## Syntax

_expression_.**InsertChartField** (_ChartFieldType_, _Formula_, _Position_)

_expression_ A variable that represents a **[TextRange2](Office.TextRange2.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ChartFieldType_|Required|[MsoChartFieldType](overview/Library-Reference/msochartfieldtype-enumeration-office.md)|Specifies the type of chart field to insert into a data label.|
| _Formula_|Optional|**String**|Specifies a cell (or range) if the **msoChartFieldFormula** constant is passed in for the _ChartFieldType_ parameter.|
| _Position_|Optional|**Integer**|Specifies the character position where the chart field is inserted. The default is to append the field to the end of the text. If the position value is out of range, the default is used.|


## Return value

TextRange2

## See also

- [TextRange2 object members](overview/Library-Reference/textrange2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]