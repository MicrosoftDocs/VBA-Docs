---
title: Chart.Export method (PowerPoint)
keywords: vbapp10.chm684028
f1_keywords:
- vbapp10.chm684028
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Export
ms.assetid: 19b95f24-c262-902e-7e96-c488affeb88d
ms.date: 08/16/2019
localization_priority: Normal
---


# Chart.Export method (PowerPoint)

Exports the chart in a graphic format.


## Syntax

_expression_.**Export** (_FileName_, _FilterName_, _Interactive_)

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the exported file.|
| _FilterName_|Optional|**Variant**|The language-independent name of the graphic filter as it appears in the registry, for example, PNG or GIF. PNG is the default if no value is supplied.|
| _Interactive_|Optional|**Variant**|**N/A**|



## Example

The following code example exports the currently selected shape as a .GIF if the shape contains a chart.

```vb
With ActiveWindow.Selection.ShapeRange(1)
    If .HasChart Then
        .Chart.Export _
            FileName:="current_sales.gif", FilterName:="GIF"
    End If
End With
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
