---
title: Application.ShowStartupDialog property (PowerPoint)
keywords: vbapp10.chm502051
f1_keywords:
- vbapp10.chm502051
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ShowStartupDialog
ms.assetid: acbd2597-c835-e285-e52c-5c86349d3199
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ShowStartupDialog property (PowerPoint)

Determines whether to display the  **New Presentation** task pane when Microsoft PowerPoint is started. Read/write.


## Syntax

_expression_. `ShowStartupDialog`

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Return value

MsoTriState


## Remarks

The value of the  **ShowStartupDialog** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|Hides the  **New Presentation** task pane.|
|**msoTrue**| The default. Displays the **New Presentation** task pane.|

## Example

The following line of code hides the  **New Presentation** task pane when PowerPoint starts.


```vb
Sub DontShowStartup()

    Application.ShowStartupDialog = msoFalse

End Sub
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]