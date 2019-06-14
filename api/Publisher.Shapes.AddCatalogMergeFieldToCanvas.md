---
title: Shapes.AddCatalogMergeFieldToCanvas method (Publisher)
keywords: vbapb10.chm2162760
f1_keywords:
- vbapb10.chm2162760
ms.prod: publisher
api_name:
- Publisher.Shapes.AddCatalogMergeFieldToCanvas
ms.assetid: 30cd45d0-97f0-ab01-31c2-8d819b435b1b
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddCatalogMergeFieldToCanvas method (Publisher)

Adds a catalog merge field of the specified type to the canvas. Returns nothing.


## Syntax

_expression_.**AddCatalogMergeFieldToCanvas** (_CanvasId_, _CatalogMergeFieldType_, _DbCol_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_CanvasId_ |Required| **Integer**|The ID of the canvas to which to add the catalog merge field.|
|_CatalogMergeFieldType_ |Required| **[PbCatalogMergeFieldType](publisher.pbcatalogmergefieldtype.md)**|The type (picture or text) of the catalog merge field to add.|
|_DbCol_ |Required| **Integer**|The number of the column in the data source that contains the catalog merge information.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]