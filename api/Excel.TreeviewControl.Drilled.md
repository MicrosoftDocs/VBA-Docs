---
title: TreeviewControl.Drilled property (Excel)
keywords: vbaxl10.chm666074
f1_keywords:
- vbaxl10.chm666074
api_name:
- Excel.TreeviewControl.Drilled
ms.assetid: 5e4f1b52-a02f-655b-f3c8-b5e7aa54d928
ms.date: 05/18/2019
ms.localizationpriority: medium
---


# TreeviewControl.Drilled property (Excel)

Sets the "drilled" (expanded or visible) status of the cube field members in the hierarchical member-selection control of a cube field. This property is used primarily for macro recording and isn't intended for any other use. Read/write.


## Syntax

_expression_.**Drilled**

_expression_ A variable that represents a **[TreeviewControl](Excel.TreeviewControl.md)** object.


## Remarks

The **Drilled** property accepts an array. Each element of the array corresponds to a level of the cube field that has been expanded. The maximum number of elements is the number of levels in the cube field. Each element of the array is an array of type **String**, containing unique member names that are visible (expanded) at the corresponding level of the control. 

To determine when members are explicitly hidden in an expanded view, see the **[Hidden](Excel.TreeviewControl.Hidden.md)** property of the **TreeviewControl** object. 

> [!NOTE] 
> This property does not return an array that represents the drilled status of the cube field members in the hierarchical member-selection control of a cube field.


## Example

This example expands the second-level members of the first cube field in the first PivotTable report on the active worksheet.

```vb
ActiveSheet.PivotTables("PivotTable1").CubeFields(1) _ 
 .TreeviewControl.Drilled = _ 
 Array(Array("", "", "", "", "", "", "", "", _ 
 "", "", "", ""), _ 
 Array("[state].[states].[AB]", _ 
 "[state].[states].[CA]", _ 
 "[state].[states].[IN]", _ 
 "[state].[states].[KS]", _ 
 "[state].[states].[KY]", _ 
 "[state].[states].[MD]", _ 
 "[state].[states].[MI]", _ 
 "[state].[states].[OH]", _ 
 "[state].[states].[OR]", _ 
 "[state].[states].[TN]", _ 
 "[state].[states].[UT]", _ 
 "[state].[states].[WA]"))
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]