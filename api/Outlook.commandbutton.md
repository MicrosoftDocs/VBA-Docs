---
title: CommandButton object (Outlook Forms Script)
keywords: olfm10.chm2000490
f1_keywords:
- olfm10.chm2000490
ms.prod: outlook
ms.assetid: bb2bcfaa-e7a5-cedc-2ed7-bcc17a4d8fb6
ms.date: 06/08/2017
localization_priority: Normal
---


# CommandButton object (Outlook Forms Script)

Starts, ends, or interrupts an action or series of actions.


## Remarks

The macro or event procedure assigned to the **CommandButton** control's **[Click](Outlook.commandbutton.click.md)** event determines what the **CommandButton** does. For example, you can create a **CommandButton** that opens another form. You can also display text, a picture, or both on a **CommandButton**.

The only event for a **CommandButton** is the **Click** event.

## Events

|Name|Description|
|:-----|:-----|
| [Click](Outlook.commandbutton.click.md)|Occurs when the user clicks inside the control.|


## Properties

|Name|Description|
|:-----|:-----|
| [Accelerator](Outlook.commandbutton.accelerator.md)|Returns or sets the accelerator key for a control. Read/write.|
| [AutoSize](Outlook.commandbutton.autosize.md)|Returns or sets a **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.|
| [BackColor](Outlook.commandbutton.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [BackStyle](Outlook.commandbutton.backstyle.md)|Returns or sets an **Integer** that specifies the background style for an object. Read/write.|
| [Caption](Outlook.commandbutton.caption.md)|Returns or sets a **String** that appears on the button to describe what it does. Read/write.|
| [Enabled](Outlook.commandbutton.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [ForeColor](Outlook.commandbutton.forecolor.md)|Returns or sets a **Long** that specifies the foreground color of an object. Read/write.|
| [Locked](Outlook.commandbutton.locked.md)|Returns or sets a **Boolean** that specifies whether a control can be edited. Read/write.|
| [MouseIcon](Outlook.commandbutton.mouseicon.md)|Returns a **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](Outlook.commandbutton.mousepointer.md)|Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [Picture](Outlook.commandbutton.picture.md)|Returns a **String** that specifies the full path name of a bitmap to display on a control. Read-only.|
| [PicturePosition](Outlook.commandbutton.pictureposition.md)|Returns or sets an **Integer** that specifies the location of the picture relative to its caption. Read/write.|
| [TakeFocusOnClick](Outlook.commandbutton.takefocusonclick.md)|Returns or sets a **Boolean** that specifies whether a control takes the focus when clicked. Read/write.|
| [WordWrap](Outlook.commandbutton.wordwrap.md)|Returns or sets a **Boolean** that specifies whether the contents of a control automatically wrap at the end of a line and the control expands to fit the text. Read/write.|





[!include[Support and feedback](~/includes/feedback-boilerplate.md)]