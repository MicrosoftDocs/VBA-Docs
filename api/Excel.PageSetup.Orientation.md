---
title: PageSetup.Orientation property (Excel)
keywords: vbaxl10.chm473090
f1_keywords:
- vbaxl10.chm473090
ms.prod: excel
api_name:
- Excel.PageSetup.Orientation
ms.assetid: 9e41d5c8-e887-3212-c298-c2921137ec9c
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.Orientation property (Excel)

Returns or sets an **[XlPageOrientation](Excel.XlPageOrientation.md)** value that represents the portrait or landscape printing mode.


## Syntax

_expression_.**Orientation**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Example

This example sets Sheet1 to be printed in landscape orientation.

```vb
Worksheets("Sheet1").PageSetup.Orientation = xlLandscape
```

<br/>

This example sets the currently active sheet to be printed in portrait orientation.

```vb
ActiveSheet.PageSetup.Orientation = xlPortrait
```

<br/>

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




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
