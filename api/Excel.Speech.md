---
title: Speech object (Excel)
keywords: vbaxl10.chm718072
f1_keywords:
- vbaxl10.chm718072
ms.prod: excel
api_name:
- Excel.Speech
ms.assetid: 1ddd61bc-989e-4766-423e-515ec5d1c23a
ms.date: 04/02/2019
localization_priority: Normal
---


# Speech object (Excel)

Contains methods and properties that pertain to speech.


## Remarks

Use the **[Speech](Excel.Application.Speech.md)** property of the **Application** object to return a **Speech** object.


## Example

After a **Speech** object is returned, you can use the **Speak** method of the **Speech** object to play back the contents of a string. In the following example, Microsoft Excel plays back "Hello". This example assumes that speech features have been installed on the host system.

> [!NOTE] 
> There is a speech feature in the setup tree that pertains to Dictation and Command & Control that does not have to be installed.

```vb
Sub UseSpeech() 
 
 Application.Speech.Speak "Hello" 
 
End Sub()
```

## Methods

- [Speak](Excel.Speech.Speak.md)

## Properties

- [Direction](Excel.Speech.Direction.md)
- [SpeakCellOnEnter](Excel.Speech.SpeakCellOnEnter.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]