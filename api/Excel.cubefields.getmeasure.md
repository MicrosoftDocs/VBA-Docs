---
title: CubeFields.GetMeasure method (Excel)
keywords: vbaxl10.chm670078
f1_keywords:
- vbaxl10.chm670078
ms.prod: excel
ms.assetid: 26647294-66df-4691-fa8e-d14cb869145b
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeFields.GetMeasure method (Excel)

Given an attribute hierarchy, returns an implicit measure for the given function that corresponds to this attribute. If an implicit measure does not exist, a new implicit measure is created and added to the **CubeFields** collection.


## Syntax

_expression_.**GetMeasure** (_AttributeHierarchy_, _Function_, _Caption_)

_expression_ A variable that represents a **[CubeFields](Excel.CubeFields.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _AttributeHierarchy_|Required|**Variant**|The unique cube field that is an attribute hierarchy (**[XlCubeFieldType](excel.xlcubefieldtype.md)** = **xlHierarchy**, and **[XlCubeFieldSubType](excel.xlcubefieldsubtype.md)** = **xlCubeAttribute**).|
| _Function_|Required| **[XlConsolidationFunction](excel.xlconsolidationfunction.md)** |The function performed in the added data field.|
| _Caption_|Optional|**Variant**|The label used in the PivotTable report to identify this measure. If the measure already exists, _Caption_ will overwrite the existing label of this measure.|

## Remarks

Getting a measure by using the **GetMeasure** function will work for these functions only: **Count**, **Sum**, **Average**, **Max**, and **Min**. 

For example, these will work: 

- `Get CubeField0 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlCount, "NumCarsOwnedCount")`

- `Set CubeField1 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlSum, "NumCarsOwnedSum")`

- `Set CubeField2 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlAverage, "NumCarsOwnedAverage")`

- `Set CubeField4 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlMax, "NumCarsOwnedMax")`

- `Set CubeField5 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlMin, "NumCarsOwnedMin")`

These will not work: 

- `Set CubeField3 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlCountNums, "NumCarsOwnedCountNums")`

- `Set CubeField6 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlProduct, "NumCarsOwnedProduct")`

- `Set CubeField7 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlStDev, "NumCarsOwnedStDev")`

- `Set CubeField8 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlStDevP, "NumCarsOwnedStDevP")`

## Return value

**CUBEFIELD**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]