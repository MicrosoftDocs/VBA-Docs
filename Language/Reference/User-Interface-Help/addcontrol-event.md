---
title: AddControl Event
keywords: fm20.chm2000010
f1_keywords:
- fm20.chm2000010
ms.prod: office
api_name:
- Office.AddControl
ms.assetid: 9febc628-1d26-9ecf-7f04-7c9431a7b9c8
ms.date: 06/08/2017
---


# AddControl Event



Occurs when a control is inserted onto a form, a  **Frame**, or a **Page** of a **MultiPage**.
<<<<<<< HEAD
 **Syntax**
=======

## Syntax

>>>>>>> master
For Frame **Private Sub**_object_ _**AddControl( )**
For MultiPage **Private Sub**_object_ _**AddControl(**_index_**As Long**, _ctrl_**As Control)**
The  **AddControl** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _index_|Required. The index of the  **Page** that will contain the new control.|
| _ctrl_|Required. The control to be added.|

<<<<<<< HEAD
 **Remarks**
The AddControl event occurs when a control is added at [run time](../../Glossary/vbe-glossary.md). This event is not initiated when you add a control at [design time](../../Glossary/vbe-glossary.md), nor is it initiated when a form is initially loaded and displayed at run time.
=======
## Remarks

The AddControl event occurs when a control is added at [run time](../../Glossary/vbe-glossary.md#run-time). This event is not initiated when you add a control at [design time](../../Glossary/vbe-glossary.md#design-time), nor is it initiated when a form is initially loaded and displayed at run time.
>>>>>>> master
The default action of this event is to add a control to the specified form,  **Frame**, or **MultiPage**.
The  **Add** method initiates the AddControl event.

