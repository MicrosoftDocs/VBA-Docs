---
title: Form.AfterLayout event (Access)
keywords: vbaac10.chm13682
f1_keywords:
- vbaac10.chm13682
ms.prod: access
api_name:
- Access.Form.AfterLayout
ms.assetid: 3b500c32-e1aa-ad06-432f-981253767c3d
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.AfterLayout event (Access)

Occurs after all charts in the specified PivotChart view have been laid out, but before they have been rendered.


## Syntax

_expression_.**AfterLayout** (_drawObject_)

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _drawObject_|Required|**Object**|A **ChChartDraw** object. Use the methods and properties of this object to draw objects on the chart.|

## Return value

Nothing


## Remarks

During this event, you can reposition the **ChTitle**, **ChLegend**, **ChChart**, and **ChAxis** objects of each PivotChart view by changing their **Left** and **Top** properties. You can reposition the **ChPlotArea** object by changing its **Left**, **Top**, **Right**, and **Bottom** properties. These properties cannot be changed outside of this event.


## Example

The following example demonstrates the syntax for a subroutine that traps the **AfterLayout** event.

```vb
Private Sub Form_AfterLayout(ByVal drawObject As Object) 
 MsgBox "The PivotChart view has been laid out." 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]