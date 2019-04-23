---
title: Procedure declaration does not match description of event or procedure having same name
keywords: vblr6.chm1040357
f1_keywords:
- vblr6.chm1040357
ms.prod: office
ms.assetid: d7b51272-3bbb-30ff-33df-202a2d89fd87
ms.date: 06/08/2017
localization_priority: Normal
---


# Procedure declaration does not match description of event or procedure having same name

Your [class module](../../Glossary/vbe-glossary.md#class-module) has a procedure name that conflicts with the name of an event. This error has the following cause and solution:

- A [procedure](../../Glossary/vbe-glossary.md#procedure) has the same name as an event, but does not have the same signature (that is, the number and types of the [parameters](../../Glossary/vbe-glossary.md#parameter)). This can occur if you do something such as add a new parameter to an event procedure. For example, if you modify the definition of a form's Form_Load event procedure as follows, this error will occur:
    
  ```vb
    Sub Form_Load (MyParam As Integer) 
    . . . 
    End Sub
  ```

  If the procedure isn't the event procedure corresponding to the event, change its name. If the procedure corresponds to the event, make the parameter list agree with that required by the event (if any).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
