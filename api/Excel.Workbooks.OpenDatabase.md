---
title: Workbooks.OpenDatabase method (Excel)
keywords: vbaxl10.chm203084
f1_keywords:
- vbaxl10.chm203084
ms.prod: excel
api_name:
- Excel.Workbooks.OpenDatabase
ms.assetid: 09f38ddb-48f9-91af-4b0f-4087c9139ab9
ms.date: 05/18/2019
localization_priority: Normal
---


# Workbooks.OpenDatabase method (Excel)

Returns a **[Workbook](Excel.Workbook.md)** object representing a database.


## Syntax

_expression_.**OpenDatabase** (_FileName_, _CommandText_, _CommandType_, _BackgroundQuery_, _ImportDataAs_)

_expression_ A variable that represents a **[Workbooks](Excel.Workbooks.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The connection string that contains the location and file name of the database.|
| _CommandText_|Optional| **Variant**|The command text of the query.|
| _CommandType_|Optional| **Variant**|The command type of the query. Specify one of the constants of the **[XlCmdType](Excel.XlCmdType.md)** enumeration: **xlCmdCube**, **xlCmdList**, **xlCmdSql**, **xlCmdTable**, and **xlCmdDefault**.|
| _BackgroundQuery_|Optional| **Variant**|This parameter is a variant data type but you can only pass a **Boolean** value. If you pass **True**, the query is performed in the background (asynchronously). The default value is **False**.|
| _ImportDataAs_|Optional| **Variant**|This parameter uses one of the values of the **[XlImportDataAs](Excel.XlImportDataAs.md)** enumeration. The two values of this enum are **xlPivotTableReport** and **xlQueryTable**. Pass one of these values to return the data as a PivotTable or QueryTable. The default value is **xlQueryTable**.|

## Return value

**Workbook**


## Example

In this example, Microsoft Excel opens the Northwind.mdb file. This example assumes that a file called Northwind.mdb exists on the C:\ drive.


```vb
Sub UseOpenDatabase() 
 ' Open the Northwind database in the background and create a PivotTable 
 Workbooks.OpenDatabase Filename:="c:\Northwind.mdb", _ 
 CommandText:="Orders", _ 
 CommandType:=xlCmdTable, _ 
 BackgroundQuery:=True, _ 
 ImportDataAs:=xlPivotTableReport 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
