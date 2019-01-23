---
title: MsoEnvelope.CommandBars property (Office)
keywords: vbaof11.chm11005
f1_keywords:
- vbaof11.chm11005
ms.prod: office
api_name:
- Office.MsoEnvelope.CommandBars
ms.assetid: ac2a7180-044a-e945-98f9-1d2fa76e7cb8
ms.date: 01/22/2019
localization_priority: Normal
---


# MsoEnvelope.CommandBars property (Office)

Gets a **CommandBars** collection. Read-only.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**CommandBars**

_expression_ A variable that represents an **[MsoEnvelope](Office.MsoEnvelope.md)** object.


## Example

The following example return the **CommandBars** collection from the **MsoEnvelope** object in Microsoft Word.


```vb
Dim cbars As CommandBars 
Set cbars = Application.ActiveDocument.MailEnvelope.Commandbars 

```


## See also

- [MsoEnvelope object members](overview/library-reference/msoenvelope-members-office.md)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]


