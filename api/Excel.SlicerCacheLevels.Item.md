---
title: SlicerCacheLevels.Item property (Excel)
keywords: vbaxl10.chm899074
f1_keywords:
- vbaxl10.chm899074
ms.prod: excel
api_name:
- Excel.SlicerCacheLevels.Item
ms.assetid: 4cf91d69-7489-9752-2b8e-ec5c7ce1a293
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerCacheLevels.Item property (Excel)

Returns the specified **[SlicerCacheLevel](Excel.SlicerCacheLevel.md)** object from the collection, or if no level is specified, returns the first **SlicerCacheLevel** object in the collection.


## Syntax

_expression_.**Item** (_Level_)

_expression_ A variable that returns a **[SlicerCacheLevels](Excel.SlicerCacheLevels.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Level_|Optional| **Variant**|The MDX unique name of the level or index number of the object.|

## Example

The following example retrieves a **SlicerCacheLevel** object that represents the Country level of the Customer Geography hierarchy from the **SlicerCacheLevel** collection of the Country slicer.

```vb
ActiveWorkbook.SlicerCaches("Slicer_Country"). _ 
 SlicerCacheLevels("[Customer].[Customer Geography].[Country]")
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]