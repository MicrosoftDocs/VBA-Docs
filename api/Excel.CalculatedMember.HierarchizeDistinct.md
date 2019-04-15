---
title: CalculatedMember.HierarchizeDistinct property (Excel)
keywords: vbaxl10.chm686083
f1_keywords:
- vbaxl10.chm686083
ms.prod: excel
api_name:
- Excel.CalculatedMember.HierarchizeDistinct
ms.assetid: 3845d280-5044-3510-38e0-51c22ba04a38
ms.date: 04/13/2019
localization_priority: Normal
---


# CalculatedMember.HierarchizeDistinct property (Excel)

Returns or sets whether to order and remove duplicates when displaying the hierarchy of the specified named set in a PivotTable report based on an OLAP cube. Read/write **Boolean**.


## Syntax

_expression_.**HierarchizeDistinct**

_expression_ A variable that returns a **[CalculatedMember](Excel.CalculatedMember.md)** object.


## Return value

**Boolean**


## Remarks

**True** if the hierarchy of the named set is displayed as ordered with duplicates removed; otherwise, **False**.

The value of this property corresponds to the **Automatically order and remove duplicates from the set** check box in the **New/Modify Set** dialog box when a named set is created or edited. 

To create a new named set from data in a PivotTable based on an OLAP data source, choose the PivotTable, choose **Field, Items, & Sets** on the **PivotTable Tools Options** tab on the ribbon, choose **Manage Sets**, choose **New** in the **Set Manager** dialog box, and then choose **Create Set using MDX**. 

This will display the **New Set** dialog box, which contains the **Automatically order and remove duplicates from the set** check box. Similarly, if you select an existing named set in the **Set Manager** dialog box, and then choose **Edit**, the **Modify Set** dialog box is displayed.

This property, along with the **[DisplayFolder](Excel.CalculatedMember.DisplayFolder.md)** and **[Dynamic](Excel.CalculatedMember.Dynamic.md)** properties, can only be read for named sets (which are represented by **CalculatedMember** objects where the **[Type](Excel.CalculatedMember.Type.md)** property equals **xlCalculatedSet**). 

These properties cannot be read for calculated members or measures (which are represented by **CalculatedMember** objects where the **Type** property equals **xlCalculatedMember**). If you attempt to read these properties for calculated members or measures, a run-time error is raised.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]