---
title: Parameter.SetParam method (Excel)
keywords: vbaxl10.chm523079
f1_keywords:
- vbaxl10.chm523079
ms.prod: excel
api_name:
- Excel.Parameter.SetParam
ms.assetid: af1f5b0a-75a1-ae85-b291-cc3ab514b0a3
ms.date: 05/03/2019
localization_priority: Normal
---


# Parameter.SetParam method (Excel)

Defines a parameter for the specified query table.


## Syntax

_expression_.**SetParam** (_Type_, _Value_)

_expression_ A variable that represents a **[Parameter](Excel.Parameter.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlParameterType](Excel.XlParameterType.md)**|One of the constants of **XlParameterType**, which specifies the parameter type.|
| _Value_|Required| **Variant**|The value of the specified parameter, as shown in the description of the _Type_ argument.|


## Example

This example changes the SQL statement for query table one. The clause `(city=?)` indicates that the query is a parameter query, and the example sets the value of city to the constant `Oakland`.

```vb
Set qt = Sheets("sheet1").QueryTables(1) 
qt.Sql = "SELECT * FROM authors WHERE (city=?)" 
Set param1 = qt.Parameters.Add("City Parameter", _ 
 xlParamTypeVarChar) 
param1.SetParam xlConstant, "Oakland" 
qt.Refresh
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]