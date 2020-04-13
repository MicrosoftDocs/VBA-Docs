---
title: Application.BoxFormat method (Project)
keywords: vbapj.chm2388
f1_keywords:
- vbapj.chm2388
ms.prod: project-server
api_name:
- Project.Application.BoxFormat
ms.assetid: bc2c0b19-c030-3063-4842-cf1bb146f73f
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BoxFormat method (Project)

Formats individual boxes in the Network Diagram view (PERT chart).


## Syntax

_expression_. `BoxFormat`( `_ProjectName_`, `_TaskID_`, `_DataTemplate_`, `_HorizontalGridlines_`, `_VerticalGridlines_`, `_BorderShape_`, `_BorderColor_`, `_BorderWidth_`, `_BackgroundColor_`, `_BackgroundPattern_`, `_Reset_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ProjectName_|Optional|**String**|The name of the project containing  **TaskID** when working with consolidated projects. The default value is the name of the active project.|
| _TaskID_|Optional|**Long**|The identification number of the task represented by the box to change. The default behavior is to change the boxes that represent one or more selected tasks.|
| _DataTemplate_|Optional|**String**|The name of the data template to use.|
| _HorizontalGridlines_|Optional|**Boolean**|**True** if horizontal gridlines separate each row in the box; otherwise, **False**.|
| _VerticalGridlines_|Optional|**Boolean**|**True** if vertical gridlines separate each column in the box; otherwise, **False**.|
| _BorderShape_|Optional|**Long**|The shape of the box border. Can be one of the **[PjBoxShape](Project.PjBoxShape.md)** constants.|
| _BorderColor_|Optional|**Long**|The color of the box border. Can be one of the **[PjColor](Project.PjColor.md)** constants.|
| _BorderWidth_|Optional|**Long**|Specifies the box border width, where values can be 1 to 4 for the four line widths shown in the **Format Box** dialog box.|
| _BackgroundColor_|Optional|**Long**|The color of the box background. Can be one of the **[PjColor](Project.PjColor.md)** constants.|
| _BackgroundPattern_|Optional|**Long**|The pattern for the background. Can be one of the **[PjBackgroundPattern](Project.PjBackgroundPattern.md)** constants.|
| _Reset_|Optional|**Boolean**|**True** if the box formatting is reset to the default style as shown in the **Box Styles** dialog box. If **Reset** is **True**, all arguments except **ProjectName** and **TaskID** are ignored.|

## Return value

 **Boolean**


## Remarks

If  **TaskID** is specified, the associated task cannot be hidden due to application of a filter or a collapsed outline structure.

Using the **BoxFormat** method without specifying any arguments displays the **Format Box** dialog box for the selected tasks. If no tasks are selected, the **BoxFormat** method has no effect.

Use the **BoxFormat** method to change the formatting of boxes from their default styles. To define the default styles, use the **BoxStylesEdit** or **BoxStylesEditEx** method.

To format Network Diagram boxes using hexadecimal values for  _BorderColor_ and _BackgroundColor_, see the **[BoxFormatEx](Project.Application.BoxFormatEx.md)** method.


## Example

The following example changes the border color to red and the background color to a light blue dithered pattern.


```vb
Sub BoxFormat_Color() 
 'Activate the Network Diagram view 
 ViewApply Name:="Network Diagram" 
 
 BoxFormat TaskID:="2", bordershape:=pjBoxRoundedRectangle, VerticalGridlines:=True, _ 
 BorderWidth:=2, backgroundpattern:=pjBackgroundLightDither, _ 
 Backgroundcolor:=pjBlue, BorderColor:=pjRed
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]