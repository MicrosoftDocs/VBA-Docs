---
title: TextRange2.InsertChartField method (PowerPoint)
ms.assetid: 42c07916-74e1-46c2-8cbc-5777c9fe1ae4
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# TextRange2.InsertChartField method (PowerPoint)

Inserts a field into the body of a data label in a chart. 

This method applies only to data labels in a chart. Calling this method on any other kind of [TextRange2](Office.TextRange2.md) object will raise a run-time error.

## Syntax

_expression_. `InsertChartField`_(ChartFieldType,_ _Formula,_ _Position)_

_expression_ A variable that represents a 'TextRange2' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ChartFieldType_|Required|[MsoChartFieldType](overview/Library-Reference/msochartfieldtype-enumeration-office.md)|Specifies the type of chart field to insert into a data label.|
| _Formula_|Optional|**String**|Specifies a cell (or range) if the  **MsoChartFieldFormula** constant is passed in for the _ChartFieldType_ parameter.|
| _Position_|Optional|**Integer**|Specifies the character position where the chart field is inserted. The default is to append the field to the end of the text. If the position value is out of range, the default is used.|
| _ChartFieldType_|Required|MSOCHARTFIELDTYPE||
| _Formula_|Optional|**String**||
| _Position_|Optional|INT||


## Return value

[TextRange2](Office.TextRange2.md)


## See also


[TextRange2 object (PowerPoint)](PowerPoint.textrange2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]