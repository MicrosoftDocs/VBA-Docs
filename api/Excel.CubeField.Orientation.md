---
title: CubeField.Orientation property (Excel)
keywords: vbaxl10.chm668077
f1_keywords:
- vbaxl10.chm668077
ms.prod: excel
api_name:
- Excel.CubeField.Orientation
ms.assetid: b134cefe-7df0-dc9f-0f7d-e93f2cb0e303
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeField.Orientation property (Excel)

Returns or sets an **[XlPivotFieldOrientation](Excel.XlPivotFieldOrientation.md)** value that represents the location of the field in the specified PivotTable report.


## Syntax

_expression_.**Orientation**

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Remarks

For OLAP data sources, setting this property for one field in a hierarchy sets the orientation for the other fields in the same hierarchy. 

Dimension fields can only be oriented in the row, column, and page field areas of the PivotTable report. Measure fields can only be oriented in the data area. 

Setting a hierarchy or data field to **xlHidden** removes the hierarchy or field from the PivotTable report.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]