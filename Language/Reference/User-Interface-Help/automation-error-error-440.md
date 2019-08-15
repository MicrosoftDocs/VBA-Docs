---
title: Automation error (Error 440)
keywords: vblr6.chm1000440
f1_keywords:
- vblr6.chm1000440
ms.prod: office
ms.assetid: 7b4be799-038b-8f70-d893-848fcfa92993
ms.date: 08/14/2019
localization_priority: Priority
---


# Automation error (Error 440)

When you access [Automation objects](../../Glossary/vbe-glossary.md#automation-object), specific types of errors can occur. This error has the following possible causes and solutions:

- An error occurred while executing a [method](../../Glossary/vbe-glossary.md#method) or getting or setting a [property](../../Glossary/vbe-glossary.md#property) of an [object variable](../../Glossary/vbe-glossary.md#object-variable). The error was reported by the application that created the object.
    
  Check the properties of the **Err** object to determine the source and nature of the error. Also try using the **On Error Resume Next** statement immediately before the accessing statement, and then check for errors immediately following the accessing statement.
    
- The Office add-in that you are trying to use has been disabled by your System Administrator.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Add-ins note](~/includes/addinsnote.md)]

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
