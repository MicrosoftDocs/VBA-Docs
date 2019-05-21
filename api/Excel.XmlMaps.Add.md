---
title: XmlMaps.Add method (Excel)
keywords: vbaxl10.chm756073
f1_keywords:
- vbaxl10.chm756073
ms.prod: excel
api_name:
- Excel.XmlMaps.Add
ms.assetid: 0197c932-73bf-024e-35b1-aba984175aee
ms.date: 05/21/2019
localization_priority: Normal
---


# XmlMaps.Add method (Excel)

Adds an XML map to the specified workbook.


## Syntax

_expression_.**Add** (_Schema_, _RootElementName_)

_expression_ An expression that returns an **[XmlMaps](Excel.XmlMaps.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Schema_|Required| **String**|The schema to be added as an XML map. The string can be a path to a schema file, or the schema itself. The path can be specified in the Universal Naming Convention (UNC) or Uniform Resource Locator (URL) format.|
| _RootElementName_|Optional| **Variant**|The name of the root element. This argument can be ignored if the schema contains only one root element.|


## Return value

An **[XmlMap](Excel.XmlMap.md)** object that represents the new XML map.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]