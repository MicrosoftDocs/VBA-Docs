---
title: PageSetup.Orientation property (Excel)
keywords: vbaxl10.chm473090
f1_keywords:
- vbaxl10.chm473090
ms.prod: excel
api_name:
- Excel.PageSetup.Orientation
ms.assetid: 9e41d5c8-e887-3212-c298-c2921137ec9c
ms.date: 06/08/2017
localization_priority: Priority
---


# PageSetup.Orientation property (Excel)

Returns or sets a  **[xlPageOrientation](Excel.XlPageOrientation.md)** value that represents the portrait or landscape printing mode.


## Syntax

_expression_. `Orientation`

_expression_ A variable that represents a [PageSetup](Excel.PageSetup.md) object.


## Example

This example sets Sheet1 to be printed in landscape orientation.

```vb
Worksheets("Sheet1").PageSetup.Orientation = xlLandscape
```
This example sets the currently active sheet to be printed in portrait orientation.

```vb
ActiveSheet.PageSetup.Orientation = xlPortrait
```

This procedure switches the orientation to the opposite option.

```vb
Sub SwitchOrientation()
    Dim ps As PageSetup
    Set ps = ActiveSheet.PageSetup

    If ps.Orientation = xlLandscape Then
        ps.Orientation = xlPortrait
    Else
        ps.Orientation = xlLandscape
    End If
End Sub
```


## See also


[PageSetup Object](Excel.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]