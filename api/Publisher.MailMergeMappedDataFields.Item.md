---
title: MailMergeMappedDataFields.Item method (Publisher)
keywords: vbapb10.chm6488064
f1_keywords:
- vbapb10.chm6488064
ms.prod: publisher
api_name:
- Publisher.MailMergeMappedDataFields.Item
ms.assetid: c1c9acde-d1e5-25d3-1b59-3e848f3881b6
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeMappedDataFields.Item method (Publisher)

Returns an individual object in a specified collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[MailMergeMappedDataFields](Publisher.MailMergeMappedDataFields.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Required| **Variant**|The number or name of the field or list box item to return.|

## Return value

MailMergeMappedDataField


## Example

This example returns the City field from a mapped data fields object.

```vb
Dim mmfTemp As MailMergeMappedDataField 
 
Set mmfTemp = ActiveDocument.MailMerge _ 
 .DataSource.MappedDataFields.Item(Index:="City")
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]