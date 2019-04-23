---
title: Global.System property (Word)
keywords: vbawd10.chm163119113
f1_keywords:
- vbawd10.chm163119113
ms.prod: word
api_name:
- Word.Global.System
ms.assetid: b1450081-e237-b45a-658e-f7c70bb0a1dc
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.System property (Word)

Returns a  **System** object, which can be used to return system-related information and perform system-related tasks.


## Syntax

_expression_. `System`

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


## Example

This example returns information about the system.


```vb
processor = System.ProcessorType 
enviro = System.OperatingSystem
```

This example establishes a connection to a network drive.




```vb
System.Connect Path:="\\Project\Info"
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]