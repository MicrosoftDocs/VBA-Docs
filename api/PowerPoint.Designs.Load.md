---
title: Designs.Load method (PowerPoint)
keywords: vbapp10.chm643005
f1_keywords:
- vbapp10.chm643005
ms.prod: powerpoint
api_name:
- PowerPoint.Designs.Load
ms.assetid: 8926e038-4b01-da8d-3e0f-6b5cdd82f1c7
ms.date: 06/08/2017
localization_priority: Normal
---


# Designs.Load method (PowerPoint)

Returns a **Design** object that represents a design loaded into the master list of the specified presentation.


## Syntax

_expression_. `Load`( `_TemplateName_`, `_Index_` )

_expression_ A variable that represents a [Designs](PowerPoint.Designs.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TemplateName_|Required|**String**|The path to the design template.|
| _Index_|Optional|**Long**|The index number of the design template in the collection of design templates. The default is -1, which means the design template is added to the end of the list of designs in the presentation.|

## Return value

Design


## Example

This example add a design template to the beginning of the collection of design templates in the active presentation. This example assumes the "artsy.pot" template is located at the specified path.


```vb
Sub LoadDesign()
    ActivePresentation.Designs.Load TemplateName:="C:\Program Files\" & _
        "Microsoft Office\Templates\Presentation Designs\Balance.pot", Index:=1
End Sub
```


## See also


[Designs Object](PowerPoint.Designs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]