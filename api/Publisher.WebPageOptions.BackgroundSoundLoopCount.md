---
title: WebPageOptions.BackgroundSoundLoopCount property (Publisher)
keywords: vbapb10.chm544776
f1_keywords:
- vbapb10.chm544776
ms.prod: publisher
api_name:
- Publisher.WebPageOptions.BackgroundSoundLoopCount
ms.assetid: 34d34a04-5fdb-3d43-9140-fcf10b420efd
ms.date: 06/18/2019
localization_priority: Normal
---


# WebPageOptions.BackgroundSoundLoopCount property (Publisher)

Returns a **Long** value that specifies the number of times the background sound attached to a webpage is played when the page is loaded in a web browser. Read-only.


## Syntax

_expression_.**BackgroundSoundLoopCount**

_expression_ A variable that represents a **[WebPageOptions](Publisher.WebPageOptions.md)** object.


## Return value

Long


## Remarks

The **[SetBackgroundSoundRepeat](Publisher.WebPageOptions.SetBackgroundSoundRepeat.md)** method can be used to specify the number of times that the background sound file is played when the page is loaded. If using the **SetBackgroundSoundRepeat** method to specify the number of times that the background file is played, the **BackgroundSoundLoopCount** property will be equal to that specified value. Note that valid values range from 1 to 999, inclusive. Attempting to set this value outside this range will result in a run-time error.

Until the **SetBackgroundSoundRepeat** method is used to change the number of times that the background sound file is played, **BackgroundSoundLoopCount** is 1.


## Example

The following example sets the background sound for page four of the active web publication to a .wav file on the local computer. If **BackgroundSoundLoopCount** is less than three, the **SetBackgroundSoundRepeat** method is used to specify that the background sound be repeated three times. The **BackgroundSoundLoopCount** property will now be three.

```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
 If .BackgroundSoundLoopCount < 3 Then 
 .SetBackgroundSoundRepeat RepeatForever:=False, RepeatTimes:=3 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]