---
title: PivotTable.CreateCubeFile method (Excel)
keywords: vbaxl10.chm235152
f1_keywords:
- vbaxl10.chm235152
ms.prod: excel
api_name:
- Excel.PivotTable.CreateCubeFile
ms.assetid: 585641a1-c708-75fd-4789-f7a254830b57
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.CreateCubeFile method (Excel)

Creates a cube file from a PivotTable report connected to an Online Analytical Processing (OLAP) data source.


## Syntax

_expression_.**CreateCubeFile** (_File_, _Measures_, _Levels_, _Members_, _Properties_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _File_|Required| **String**|The name of the cube file to be created. It will overwrite the file if it already exists.|
| _Measures_|Optional| **Variant**|An array of unique names of measures that are to be part of the slice.|
| _Levels_|Optional| **Variant**|An array of strings. Each array item is a unique level name. It represents the lowest level of a hierarchy that is in the slice.|
| _Members_|Optional| **Variant**|An array of string arrays. The elements correspond, in order, to the hierarchies represented in the _Levels_ array. Each element is an array of string arrays that consists of the unique names of the top level members in the dimension that are to be included in the slice.|
| _Properties_|Optional| **Variant**| **False** results in no member properties being included in the slice. The default value is **True**.|

## Return value

String


## Example

This example creates a cube file titled CustomCubeFile on drive C:\ with no member properties to be included in the slice. With the _Measures_, _Levels_, and _Members_ arguments omitted from this example, the cube file will end up matching the view of the PivotTable report. This example assumes that a PivotTable report connected to an OLAP data source exists on the active worksheet.

```vb
Sub UseCreateCubeFile() 
 
 ActiveSheet.PivotTables(1).CreateCubeFile _ 
 File:="C:\CustomCubeFile", Properties:=False 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]