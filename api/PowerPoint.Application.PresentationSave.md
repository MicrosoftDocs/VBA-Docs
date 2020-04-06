---
title: Application.PresentationSave event (PowerPoint)
keywords: vbapp10.chm621005
f1_keywords:
- vbapp10.chm621005
ms.prod: powerpoint
api_name:
- PowerPoint.Application.PresentationSave
ms.assetid: 229a02a7-58e4-2445-3bd5-963e88438d7e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PresentationSave event (PowerPoint)

Occurs before any open presentation is saved.


## Syntax

_expression_. `PresentationSave`( `_Pres_` )

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The presentation to be saved.|

## Remarks

For information about using events with the  **Application** object, see [How to: Use Events with the Application Object](../powerpoint/How-to/use-events-with-the-application-object.md).


## Example

This example saves the current presentation as an HTML version 4.0 file with the name "mallard.htm." It then displays a message indicating that the current named presentation is being saved in both PowerPoint and HTML formats.


```vb
Private Sub App_PresentationSave(ByVal Pres As Presentation)
    With Pres.PublishObjects(1)

        PresName = .SlideShowName
        .SourceType = ppPublishAll
        .FileName = "C:\HTMLPres\mallard.htm"
        .HTMLVersion = ppHTMLVersion4

        MsgBox ("Saving presentation " & "'" _
            & PresName & "'" & " in PowerPoint" _
            & Chr(10) & Chr(13) _
            & " format and HTML version 4.0 format")
			
        .Publish
    End With
End Sub
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]