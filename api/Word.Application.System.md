---
title: Application.System property (Word)
keywords: vbawd10.chm158334985
f1_keywords:
- vbawd10.chm158334985
ms.prod: word
api_name:
- Word.Application.System
ms.assetid: 871f3821-4e17-1c63-9b4b-1d4e2bfc97d5
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.System property (Word)

Returns a  **[System](Word.System.md)** object, which can be used to return system-related information and perform system-related tasks.


## Syntax

_expression_. `System`

 _expression_ An expression that returns an **[Application](Word.Application.md)** object. 


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


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]