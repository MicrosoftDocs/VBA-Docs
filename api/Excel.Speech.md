---
title: Speech object (Excel)
keywords: vbaxl10.chm718072
f1_keywords:
- vbaxl10.chm718072
ms.prod: excel
api_name:
- Excel.Speech
ms.assetid: 1ddd61bc-989e-4766-423e-515ec5d1c23a
ms.date: 06/08/2017
localization_priority: Normal
---


# Speech object (Excel)

Contains methods and properties that pertain to speech.


## Remarks

Use the **[Speech](Excel.Application.Speech.md)** property of the **[Application](Excel.Application(object).md)** object to return a **[Speech](Excel.Speech.md)** object.


## Example

Once a **Speech** object is returned, you can use the **[Speak](Excel.Speech.Speak.md)** method of **Speech** object to play back the contents of a string. In the following example, Microsoft Excel plays back "Hello". This example assumes speech features have been installed on the host system.


 **Note**  There is a speech feature in the setup tree that pertains to Dictation and Command & Control that does not have to be installed.


```vb
Sub UseSpeech() 
 
 Application.Speech.Speak "Hello" 
 
End Sub()
```


## See also



[Excel Object Model Reference](./overview/Excel/object-model.md)

