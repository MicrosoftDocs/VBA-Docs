---
title: Application.ResourceAddressBook method (Project)
keywords: vbapj.chm2115
f1_keywords:
- vbapj.chm2115
ms.prod: project-server
api_name:
- Project.Application.ResourceAddressBook
ms.assetid: 012ba9fe-f86e-4d1c-ab24-7a500d8f3b0a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ResourceAddressBook method (Project)

Displays a MAPI-compliant address book from which the user can select resources for the project. 


## Syntax

_expression_. `ResourceAddressBook`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The **ResourceAddressBook** method is available only in resource views. If no email profile is available, Project displays a message that explains how to create a profile.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]