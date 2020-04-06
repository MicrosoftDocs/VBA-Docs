---
title: Application.PresentationOpen event (PowerPoint)
keywords: vbapp10.chm621006
f1_keywords:
- vbapp10.chm621006
ms.prod: powerpoint
api_name:
- PowerPoint.Application.PresentationOpen
ms.assetid: 1739cee9-cfc1-0650-de24-be699bafe910
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PresentationOpen event (PowerPoint)

Occurs after an existing presentation is opened, as it is added to the  **[Presentations](PowerPoint.Presentations.md)** collection.


## Syntax

_expression_. `PresentationOpen`( `_Pres_` )

 _expression_ An expression that returns an **[Application](PowerPoint.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The presentation that is opened.|

## Remarks

For information about using events with the  **Application** object, see [How to: Use Events with the Application Object](../powerpoint/How-to/use-events-with-the-application-object.md).

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this event maps to the following types:


- **Microsoft.Office.Interop.PowerPoint.EApplication_PresentationOpenEventHandler** (the **PresentationOpen** delegate.)
    
- **Microsoft.Office.Interop.PowerPoint.EApplication_Event.PresentationOpen** (the **PresentationOpen** event.)
    

## Example

This example modifies the background color for color scheme three, applies the modified color scheme to the presentation that was just opened, and displays the presentation in slide view.


```vb
Private Sub App_PresentationOpen(ByVal Pres As Presentation) 
    With Pres 
        Set CS3 = .ColorSchemes(3) 
        CS3.Colors(ppBackground).RGB = RGB(240, 115, 100) 
        With Windows(1) 
            .Selection.SlideRange.ColorScheme = CS3 
            .ViewType = ppViewSlide 
        End With 
    End With 
End Sub
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]