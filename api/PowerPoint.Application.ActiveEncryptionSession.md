---
title: Application.ActiveEncryptionSession property (PowerPoint)
keywords: vbapp10.chm502059
f1_keywords:
- vbapp10.chm502059
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ActiveEncryptionSession
ms.assetid: 73a174d5-a088-97d0-5f71-931456493224
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ActiveEncryptionSession property (PowerPoint)

Represents the encryption session associated with the active presentation. Read-only.


## Syntax

_expression_. `ActiveEncryptionSession`

 _expression_ An expression that returns an **[Application](PowerPoint.Application.md)** object.


## Return value

Long


## Remarks

The encryption provider mechanism manages each file on a separate process, so each file is associated with a separate encryption session.


> [!NOTE] 
> This property applies only when a presentation implements custom encryption.


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]