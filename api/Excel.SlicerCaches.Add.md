---
title: SlicerCaches.Add Method (Excel)
keywords: vbaxl10.chm895078
f1_keywords:
- vbaxl10.chm895078
ms.prod: excel
api_name:
- Excel.SlicerCaches.Add
ms.assetid: 8d6f1099-e1ea-d157-8e64-1a9956b77c1b
ms.date: 06/08/2017
---


# SlicerCaches.Add Method (Excel)

Adds a new  **[SlicerCache](Excel.SlicerCache.md)** object to the collection.


## Syntax

 _expression_. `Add`( `_Source_` , `_SourceField_` , `_Name_` , `_SlicerCacheType_` )

 _expression_ A variable that represents a '[SlicerCaches](Excel.SlicerCaches.md)' collection.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **Variant**|The data source that the new  **SlicerCache** will be based on. The argument passed to the _Source_ parameter can be a **[WorkbookConnection](Excel.WorkbookConnection.md)** object, a **[PivotTable](Excel.PivotTable.md)** object, or a string. If a **PivotTable** object is passed, the associated **[PivotCache](Excel.PivotCache.md)** object is used as the data source. If a string is passed, it is interpreted as the name of a **WorkbookConnection** object, and if no such **WorkbookConnection** object exists, a run-time error is generated.|
| _SourceField_|Required| **Variant**|The name of the field in the data source to filter by. For non-OLAP data sources, use the  **[PivotField](Excel.PivotField.md)** object from the **PivotCache** object that the slicer is based on, or the unique name of that object (the value of the **PivotField** . **[Name](Excel.PivotField.Name.md)** property). For OLAP data sources, use the MDX unique name of the hierarchy that the **SlicerCache** is based on. You can also specify a level of the OLAP hierarchy, and Excel will use the corresponding hierarchy.|
| _Name_|Optional| **Variant**|The name Excel uses to reference the slicer cache (the value of the  **SlicerCache** . **[Name](Excel.SlicerCache.Name.md)** property). If omitted, Excel will generate a name. By default, Excel concatenates "Slicer_" with the value of the **PivotField** . **[Caption](Excel.PivotField.Caption.md)** property for slicers with non-OLAP data sources, or with the value of the **CubeField** . **[Caption](Excel.CubeField.Caption.md)** property for slicers with OLAP data sources. (Replacing any spaces with "_".) If required to make the name unique in the workbook namespace, Excel adds an integer to the end of the generated name. If you specify a name that already exists in the workbook namespace, the **Add** method will fail.|
| _SlicerCacheType_|Optional|[XlSlicerCacheType](Excel.xlslicercachetype.md)|Designates the type of slicer or slicer cache.|

### Return value

SlicerCache


## Example

The following code example adds a slicer cache based on the Customer Geography OLAP hierarchy.


```vb
 ActiveWorkbook.SlicerCaches.Add(ActiveCell.PivotTable, _ 
 "[Customer].[Customer Geography]")
```


## See also


[SlicerCaches Object](Excel.SlicerCaches.md)

