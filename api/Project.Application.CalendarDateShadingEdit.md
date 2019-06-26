---
title: Application.CalendarDateShadingEdit method (Project)
keywords: vbapj.chm2343
f1_keywords:
- vbapj.chm2343
ms.prod: project-server
api_name:
- Project.Application.CalendarDateShadingEdit
ms.assetid: 73c8875c-fc54-ae8a-55de-f2640ac4c23a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CalendarDateShadingEdit method (Project)

Changes the background color and pattern of date boxes in the Calendar view.


## Syntax

_expression_. `CalendarDateShadingEdit`( `_Item_`, `_Pattern_`, `_Color_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required|**Long**|The type of calendar day to change. Can be one of the  **[PjCalendarShading](Project.PjCalendarShading.md)** constants.|
| _Pattern_|Optional|**Long**|The pattern for the type of date box specified by  **Item**. Can be one of the **[PjFillPattern](Project.PjFillPattern.md)** constants.|
| _Color_|Optional|**Long**|The color for the type of date box specified by  **Item**. Can be one of the **[PjColor](Project.PjColor.md)** constants.|

## Return value

 **Boolean**


## Remarks

Besides  _Item_, **CalendarDateShadingEdit** requires either the _Pattern_ or _Color_ parameter, or both, to run without an error. For example, the following line in the **Immediate** pane of the VBE works correctly.


```vb
? CalendarDateShadingEdit (PjCalendarShading.pjBaseWorking, , &H01dddd)
```

To edit calendar date boxes where the colors can be RGB values, use the  **[CalendarDateShadingEditEx](Project.Application.CalendarDateShadingEditEx.md)** method.


## Example

The following example changes the background color of working days in the base calendar to a stippled purple and the color of nonworking days to gray.


```vb
Sub CalendarDate_ShadingEdit() 
 ' Activate the Caldender view. 
 ViewApply Name:="Calendar" 
 
 CalendarDateShadingEdit Item:=pjBaseWorking, Pattern:=pjLightFillPattern, Color:=pjPurple 
 CalendarDateShadingEdit Item:=pjBaseNonworking, Color:=pjGray 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]