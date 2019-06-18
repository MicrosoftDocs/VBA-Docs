---
title: WebPageOptions.BackgroundSoundLoopForever property (Publisher)
keywords: vbapb10.chm544775
f1_keywords:
- vbapb10.chm544775
ms.prod: publisher
api_name:
- Publisher.WebPageOptions.BackgroundSoundLoopForever
ms.assetid: f2e90665-09e9-5215-59e4-f93e4469d0df
ms.date: 06/18/2019
localization_priority: Normal
---


# WebPageOptions.BackgroundSoundLoopForever property (Publisher)

Returns a **Boolean** value that specifies whether the background sound attached to the webpage should be repeated infinitely. Read-only.


## Syntax

_expression_.**BackgroundSoundLoopForever**

_expression_ A variable that represents a **[WebPageOptions](Publisher.WebPageOptions.md)** object.


## Return value

Boolean


## Remarks

The **[SetBackgroundSoundRepeat](Publisher.WebPageOptions.SetBackgroundSoundRepeat.md)** method is used to specify whether the background sound should be repeated infinitely after the page is loaded. Until the **SetBackgroundSoundRepeat** method is used to specify whether the background sound should be played infinitely, **BackgroundSoundLoopForever** is **False**.


## Example

The following example sets the background sound for page four of the active web publication to a .wav file on the local computer. If **BackgroundSoundLoopForever** is **False**, the **SetBackgroundSoundRepeat** method is used to specify that the background sound should be repeated infinitely. The **BackgroundSoundLoopForever** property will now be **True**.

```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
 If .BackgroundSoundLoopForever = False Then 
 .SetBackgroundSoundRepeat RepeatForever:=True 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]