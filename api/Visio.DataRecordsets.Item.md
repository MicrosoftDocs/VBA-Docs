---
title: DataRecordsets.Item property (Visio)
keywords: vis_sdr.chm16313765
f1_keywords:
- vis_sdr.chm16313765
ms.prod: visio
api_name:
- Visio.DataRecordsets.Item
ms.assetid: 8a289fb1-8cc5-eb76-efb1-c01f73c6340a
ms.date: 06/08/2017
localization_priority: Normal
---


# DataRecordsets.Item property (Visio)

Returns the **DataRecordset** object at the specified index position in the **DataRecordsets** collection. Read-only.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[DataRecordsets](Visio.DataRecordsets.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index of the object to retrieve.|

## Return value

DataRecordset


## Remarks

**Item** is the default property of the **DataRecordsets** collection.

When you retrieve objects from a collection, you can omit **Item** from the expression because it is the default property of all collections. The following statement is equivalent to the syntax example given above:

```vb
objectReturned = expression(Index)
```

 The **DataRecordsets** collection is indexed starting with 1.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]