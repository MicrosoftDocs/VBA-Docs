---
title: CalculatedMember.DisplayFolder property (Excel)
keywords: vbaxl10.chm686082
f1_keywords:
- vbaxl10.chm686082
ms.prod: excel
api_name:
- Excel.CalculatedMember.DisplayFolder
ms.assetid: 9ece45d1-4d27-0305-1189-15c414353607
ms.date: 04/13/2019
localization_priority: Normal
---


# CalculatedMember.DisplayFolder property (Excel)

Returns the display folder name for a named set. Read-only.


## Syntax

_expression_.**DisplayFolder**

_expression_ A variable that returns a **[CalculatedMember](Excel.CalculatedMember.md)** object.


## Return value

**String**


## Remarks

The value of this property corresponds to the optional value that can be entered in the **Display folder** text box of the **New/Modify Set** dialog box when a named set is created or edited. 

To create a new named set from data in a PivotTable based on an OLAP data source, choose the PivotTable, choose **Field, Items, & Sets** on the **PivotTable Tools Options** tab on the ribbon, choose **Manage Sets**, choose **New** in the **Set Manager** dialog box, and then choose **Create Set using MDX**. 

This will display the **New Set** dialog box, which contains the **Display folder** text box. Similarly, if you select an existing named set in the **Set Manager** dialog box, and then choose **Edit**, the **Modify Set** dialog box is displayed.

This property, along with the **[Dynamic](Excel.CalculatedMember.Dynamic.md)** and **[HierarchizeDistinct](Excel.CalculatedMember.HierarchizeDistinct.md)** properties, can only be read for named sets (which are represented by **CalculatedMember** objects where the **[Type](Excel.CalculatedMember.Type.md)** property equals **xlCalculatedSet**). 

These properties cannot be read for calculated members or measures (which are represented by **CalculatedMember** objects where the **Type** property equals **xlCalculatedMember**). If you attempt to read these properties for calculated members or measures, a run-time error is raised.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]