---
title: Application.DisplayGridLines property (PowerPoint)
keywords: vbapp10.chm502047
f1_keywords:
- vbapp10.chm502047
ms.prod: powerpoint
api_name:
- PowerPoint.Application.DisplayGridLines
ms.assetid: b639cd4f-26d4-4f63-2fe0-18807bdeefa5
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DisplayGridLines property (PowerPoint)

Determines whether to display gridlines in Microsoft PowerPoint. Read/write.


## Syntax

_expression_. `DisplayGridLines`

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Return value

MsoTriState


## Remarks

The value returned by the  **DisplayGridLines** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|Do not display gridlines.|
|**msoTrue**| Display gridlines.|

## Example

This example switches the display of the gridlines in PowerPoint.


```vb
Sub ToggleGridLines()

    With Application

        If .DisplayGridLines = msoTrue Then

            .DisplayGridLines = msoFalse

        Else

            .DisplayGridLines = msoTrue

        End If

    End With

End Sub
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]