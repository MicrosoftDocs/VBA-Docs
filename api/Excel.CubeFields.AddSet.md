---
title: CubeFields.AddSet method (Excel)
keywords: vbaxl10.chm670077
f1_keywords:
- vbaxl10.chm670077
ms.prod: excel
api_name:
- Excel.CubeFields.AddSet
ms.assetid: 2f40d4f3-56fc-4d98-b214-623885dc26d6
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeFields.AddSet method (Excel)

Adds a new **[CubeField](Excel.CubeField.md)** object to the **CubeFields** collection. The **CubeField** object corresponds to a set defined on the Online Analytical Processing (OLAP) provider for the cube.


## Syntax

_expression_.**AddSet** (_Name_, _Caption_)

_expression_ A variable that represents a **[CubeFields](Excel.CubeFields.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|A valid name in the SETS schema rowset.|
| _Caption_|Required| **String**|A string representing the field that will be displayed in the PivotTable view.|

## Return value

CubeField


## Remarks

If a set with the name given in the argument _Name_ does not exist, the **AddSet** method will return a run-time error.


## Example

In this example, Microsoft Excel adds a set titled My Set to the **CubeField** object. This example assumes that an OLAP PivotTable report exists on the active worksheet, and that a field titled Product exists.

```vb
Sub UseAddSet() 
 
 Dim pvtOne As PivotTable 
 Dim strAdd As String 
 Dim strFormula As String 
 Dim cbfOne As CubeField 
 
 Set pvtOne = Sheet1.PivotTables(1) 
 
 strAdd = "[MySet]" 
 strFormula = "'{[Product].[All Products].[Food].children}'" 
 
 ' Establish connection with data source if necessary. 
 If Not pvtOne.PivotCache.IsConnected Then pvtOne.PivotCache.MakeConnection 
 
 ' Add a calculated member titled "[MySet]" 
 pvtOne.CalculatedMembers.Add Name:=strAdd, _ 
 Formula:=strFormula, Type:=xlCalculatedSet 
 
 ' Add a set to the CubeField object. 
 Set cbfOne = pvtOne.CubeFields.AddSet(Name:="[MySet]", _ 
 Caption:="My Set") 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]