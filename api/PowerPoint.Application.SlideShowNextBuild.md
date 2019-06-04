---
title: Application.SlideShowNextBuild event (PowerPoint)
keywords: vbapp10.chm621012
f1_keywords:
- vbapp10.chm621012
ms.prod: powerpoint
api_name:
- PowerPoint.Application.SlideShowNextBuild
ms.assetid: 63919ea5-57e4-853a-0e5a-94e1126cbfbf
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SlideShowNextBuild event (PowerPoint)

Occurs upon mouse-click or timing animation, but before the animated object becomes visible. .


## Syntax

_expression_. `SlideShowNextBuild`(_Wn_)

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wn_|Required|**SlideShowWindow**|The active slide show window.|

## Remarks

For information about using events with the  **Application** object, see [How to: Use Events with the Application Object](../powerpoint/How-to/use-events-with-the-application-object.md).


## Example

If the current shape on slide one is a movie, this example plays the movie continuously until stopped manually by the presenter. This code is designed to be used with the second  **SlideShowNextSlide** event example.


```vb
Private Sub App_SlideShowNextBuild(ByVal Wn As SlideShowWindow)

    If EvtCounter <> 0 Then

        With ActivePresentation.Slides(1) _
                .Shapes(shpAnimArray(2, EvtCounter))

            If .Type =msoMedia Then

                If .MediaType = ppMediaTypeMovie

                    .AnimationSettings.PlaySettings _
                        .LoopUntilStopped

                End If

            End If

        End With

    End If

	EvtCounter = EvtCounter + 1

End Sub

	
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]