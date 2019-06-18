---
title: WebPageOptions.SetBackgroundSoundRepeat method (Publisher)
keywords: vbapb10.chm544777
f1_keywords:
- vbapb10.chm544777
ms.prod: publisher
api_name:
- Publisher.WebPageOptions.SetBackgroundSoundRepeat
ms.assetid: a699fa92-a36a-6722-431d-a0ce8413cfcf
ms.date: 06/18/2019
localization_priority: Normal
---


# WebPageOptions.SetBackgroundSoundRepeat method (Publisher)

Specifies whether the background sound attached to a webpage should be played infinitely after the page is loaded in a web browser, and if it should not, optionally specifies the number of times that the background sound should be played.


## Syntax

_expression_.**SetBackgroundSoundRepeat** (_RepeatForever_, _RepeatTimes_)

_expression_ A variable that represents a **[WebPageOptions](Publisher.WebPageOptions.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_RepeatForever_|Required| **Boolean**|Specifies whether the background sound should be played infinitely. The value of this parameter is used to populate the value of the **[BackgroundSoundLoopForever](Publisher.WebPageOptions.BackgroundSoundLoopForever.md)** property.|
|_RepeatTimes_|Optional| **Long**|Specifies how many times the background sound should be played. The value of this parameter is used to populate the value of the **[BackgroundSoundLoopCount](Publisher.WebPageOptions.BackgroundSoundLoopCount.md)** property.|

## Remarks

If the _RepeatForever_ parameter is set to **True**, the optional _RepeatTimes_ parameter cannot be specified. Attempting to specify _RepeatTimes_ if _RepeatForever_ is **True** results in a run-time error.

If the _RepeatForever_ parameter is set to **False**, the optional _RepeatTimes_ parameter must be specified. Omitting _RepeatTimes_ if _RepeatForever_ is **False** results in a run-time error.


## Example

The following example sets the background sound for page four of the active web publication to a .wav file on the local computer. If **BackgroundSoundLoopForever** is **False**, the **SetBackgroundSoundRepeat** method is used to specify that the background sound be repeated infinitely (note the omission of the _RepeatTimes_ parameter). 

If **BackgroundSoundLoopForever** is **True**, the **SetBackgroundSoundRepeat** method is used to specify that the background sound not be repeated infinitely, but that it should be repeated twice.

```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
 If .BackgroundSoundLoopForever = False Then 
 .SetBackgroundSoundRepeat RepeatForever:=True 
 ElseIf .BackgroundSoundLoopForever = True Then 
 .SetBackgroundSoundRepeat RepeatForever:=False, RepeatTimes:=2 
 End If 
 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]