---
title: Enter, Exit events
keywords: fm20.chm2000160
f1_keywords:
- fm20.chm2000160
ms.prod: office
ms.assetid: 4dc74a16-eead-48e5-2031-eaf5730bd857
ms.date: 11/15/2018
localization_priority: Normal
---


# Enter, Exit events

Enter occurs before a control actually receives the [focus](../../Glossary/vbe-glossary.md#focus) from a control on the same form. Exit occurs immediately before a control loses the focus to another control on the same form.

## Syntax

**Private Sub**_object_ _**Enter( )** <br/>
**Private Sub**_object_ _**Exit( ByVal**_Cancel_**As MSForms.ReturnBoolean)**

The **Enter** and **Exit** event syntaxes have these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _Cancel_|Required. Event status. **False** indicates that the control should handle the event (default). **True** indicates that the application handles the event and the focus should remain at the current control.|

## Remarks

The Enter and Exit events are similar to the GotFocus and LostFocus events in Visual Basic. Unlike GotFocus and LostFocus, the Enter and Exit events don't occur when a form receives or loses the focus.

For example, suppose you select the check box that initiates the Enter event. If you then select another control in the same form, the Exit event is initiated for the check box (because focus is moving to a different object in the same form), and then the Enter event occurs for the second control on the form.

Because the Enter event occurs before the focus moves to a particular control, you can use an Enter event procedure to display instructions; for example, you could use a macro or event procedure to display a small form or message box identifying the type of data the control typically contains.

> [!NOTE] 
> To prevent the control from losing focus, assign **True** to the _Cancel_ argument of the Exit event.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]