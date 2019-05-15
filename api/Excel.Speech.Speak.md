---
title: Speech.Speak method (Excel)
keywords: vbaxl10.chm718073
f1_keywords:
- vbaxl10.chm718073
ms.prod: excel
api_name:
- Excel.Speech.Speak
ms.assetid: d17dcf63-c837-a5b5-8267-44767b38700a
ms.date: 05/16/2019
localization_priority: Normal
---


# Speech.Speak method (Excel)

Microsoft Excel plays back the text string that is passed as an argument.


## Syntax

_expression_.**Speak** (_Text_, _SpeakAsync_, _SpeakXML_, _Purge_)

_expression_ A variable that represents a **[Speech](Excel.Speech.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text to be spoken.|
| _SpeakAsync_|Optional| **Variant**| **True** causes the text to be spoken asynchronously (the method will not wait for the text to be spoken). **False** causes the text to be spoken synchronously (the method waits for the text to be spoken before continuing). The default is **False**.|
| _SpeakXML_|Optional| **Variant**| **True** causes the text to be interpreted as XML. **False** causes the text to not be interpreted as XML, so any XML tags are read and not interpreted. The default is **False**.|
| _Purge_|Optional| **Variant**| **True** causes current speech to be terminated and any buffered text to be purged before text is spoken. **False** does not cause the current speech to be terminated and does not purge the buffered text before text is spoken. The default is **False**.|


## Example

In this example, Microsoft Excel speaks "Hello".

```vb
Sub UseSpeech() 
 
 Application.Speech.Speak "Hello" 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
