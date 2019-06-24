---
title: AccelItem.AddOnArgs property (Visio)
keywords: vis_sdr.chm14513045
f1_keywords:
- vis_sdr.chm14513045
ms.prod: visio
api_name:
- Visio.AccelItem.AddOnArgs
ms.assetid: ebc91b1e-7780-1cdd-04dc-4a859c8929ff
ms.date: 06/24/2019
localization_priority: Normal
---


# AccelItem.AddOnArgs property (Visio)

Gets or sets the argument string that you send to the add-on associated with a particular accelerator key. Read/write.


## Syntax

_expression_.**AddOnArgs**

_expression_ An expression that returns an **[AccelItem](Visio.AccelItem.md)** object.


## Return value

String


## Remarks

An argument's string can be anything appropriate for the add-on. However, the arguments are packaged together with other information into a command string, which cannot exceed 127 characters. For best results, limit arguments to 50 characters.

An object's **AddOnName** property indicates the name of the add-on to which the arguments are sent.

> [!NOTE] 
> Beginning with Visio 2002, the **AddOnName** property cannot execute a string that contains arbitrary VBA code. To call code that in previous versions of Visio you would have passed to the **AddOnName** property, move the code to a procedure in a document's VBA project that is called from the **AddOnName** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]