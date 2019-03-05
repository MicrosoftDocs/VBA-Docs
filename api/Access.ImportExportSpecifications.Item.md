---
title: ImportExportSpecifications.Item property (Access)
keywords: vbaac10.chm13340
f1_keywords:
- vbaac10.chm13340
ms.prod: access
api_name:
- Access.ImportExportSpecifications.Item
ms.assetid: 0068db82-cffb-c429-8d91-43c34a916d76
ms.date: 03/06/2019
localization_priority: Normal
---


# ImportExportSpecifications.Item property (Access)

The **Item** property returns a specific member of a collection either by position or by index. Read-only **ImportExportSpecification**.


## Syntax

_expression_.**Item** (Index)

_expression_ A variable that represents an **[ImportExportSpecifications](Access.ImportExportSpecifications.md)** object.


### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|An expression that specifies the position of a member of the collection referred to by the _expression_ argument.<br/><br/>If a numeric expression, the _Index_ argument must be a number from 0 to the value of the collection's **Count** property minus 1.<br/><br/>If a string expression, the _Index_ argument must be the name of a member of the collection.|




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]