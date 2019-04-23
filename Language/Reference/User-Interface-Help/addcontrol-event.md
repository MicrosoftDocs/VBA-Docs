---
title: AddControl event
keywords: fm20.chm2000010
f1_keywords:
- fm20.chm2000010
ms.prod: office
api_name:
- Office.AddControl
ms.assetid: 9febc628-1d26-9ecf-7f04-7c9431a7b9c8
ms.date: 11/15/2018
localization_priority: Normal
---


# AddControl event

Occurs when a control is inserted onto a form, a **[Frame](frame-control.md)**, or a **Page** of a **[MultiPage](multipage-control.md)**.

## Syntax

For Frame  <br/>
**Private Sub**_object_ _**AddControl( )**

For MultiPage  <br/>
**Private Sub**_object_ _**AddControl(**_index_**As Long**, _ctrl_**As Control)**

The **AddControl** event syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _index_|Required. The index of the **Page** that will contain the new control.|
| _ctrl_|Required. The control to be added.|

## Remarks

The AddControl event occurs when a control is added at [run time](../../Glossary/vbe-glossary.md#run-time). This event is not initiated when you add a control at [design time](../../Glossary/vbe-glossary.md#design-time), nor is it initiated when a form is initially loaded and displayed at run time.

The default action of this event is to add a control to the specified form, **Frame**, or **MultiPage**.

The **Add** method initiates the AddControl event.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]