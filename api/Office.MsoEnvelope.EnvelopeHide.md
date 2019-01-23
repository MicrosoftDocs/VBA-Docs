---
title: MsoEnvelope.EnvelopeHide event (Office)
keywords: vbaof11.chm246002
f1_keywords:
- vbaof11.chm246002
ms.prod: office
api_name:
- Office.MsoEnvelope.EnvelopeHide
ms.assetid: 066b0ed0-bd5d-f37e-6c04-66e0a59733d4
ms.date: 01/22/2019
localization_priority: Normal
---


# MsoEnvelope.EnvelopeHide event (Office)

Occurs when the user interface (UI) that corresponds to the **MsoEnvelope** object is hidden.


## Syntax

_expression_.**EnvelopeHide**

_expression_ An expression that returns an **[MsoEnvelope](Office.MsoEnvelope.md)** object.


## Remarks

The **MsoEnvelope** object provides access to functionality that lets you send documents as email messages directly from Microsoft Office applications.


## Example

The following example sets up event-handling routines for the **MsoEnvelope** object.


```vb
Public WithEvents env As MsoEnvelope 
 
Private Sub Class_Initialize() 
 Set env = Application.ActiveDocument.MailEnvelope 
End Sub 
 
Private Sub env_EnvelopeShow() 
 MsgBox "The MsoEnvelope UI is showing." 
End Sub 
 
Private Sub env_EnvelopeHide() 
 MsgBox "The MsoEnvelope UI is hidden." 
End Sub 

```


## See also

- [MsoEnvelope object members](overview/library-reference/msoenvelope-members-office.md)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]


