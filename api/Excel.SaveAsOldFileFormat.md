---
title: SaveAsOldFileFormat Method (Excel Graph)
keywords: vbagr10.chm5207946
f1_keywords:
- vbagr10.chm5207946
ms.prod: excel
api_name:
- Excel.SaveAsOldFileFormat
ms.assetid: 0fcdaf08-df42-6d0c-702b-4bd522ab0795
ms.date: 06/08/2017
localization_priority: Normal
---


# SaveAsOldFileFormat Method (Excel Graph)

In a host application such as Microsoft PowerPoint, saves a chart in the specified older file format.

_expression_. `SaveAsOldFileFormat( _MajorVersion_`,  `_MinorVersion)_`

 _expression_ Required. An expression that returns an [Application](Excel.Application-graph-property.md) object.

 **MajorVersion** Optional **Variant**. Specifies the major version number of the file format you want to use.
 **MinorVersion** Optional **Variant**. Specifies the minor version number of the file format you want to use.

## Example

This example saves the chart in Graph version 5.0 file format.


```vb
myChart.Application.SaveAsOldFileFormat MajorVersion:=5
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]