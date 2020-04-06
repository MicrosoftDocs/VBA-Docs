---
title: Frame Members (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 2af42a27-046d-4b5c-8112-0350fcffb7d6
ms.date: 06/08/2017
localization_priority: Normal
---


# Frame Members (Outlook Forms Script)

Represents a functional and visual control group.


## Methods


|Name|Description|
|:-----|:-----|
| [Copy](Outlook.Frame.copy.md)|Copies the contents of an object to the Clipboard.|
| [Cut](Outlook.Frame.cut.md)|Removes selected information from an object and transfers it to the Clipboard.|
| [Paste](Outlook.Frame.paste.md)|Transfers the contents of the Clipboard to an object.|
| [RedoAction](Outlook.Frame.redoaction.md)|Reverses the effect of the most recent  **Undo** action.|
| [Repaint](Outlook.Frame.repaint.md)|Updates the display by redrawing the frame.|
| [Scroll](Outlook.Frame.scroll.md)|Moves the scroll bar on an object.|
| [SetDefaultTabOrder](Outlook.Frame.setdefaulttaborder.md)|Sets the  **TabIndex** property of each control on a frame or page, using a default top-to-bottom, left-to-right tab order.|
| [UndoAction](Outlook.Frame.undoaction.md)|Reverses the most recent action that supports the  **Undo** command.|


## Properties


|Name|Description|
|:-----|:-----|

| [ActiveControl](Outlook.Frame.activecontrol.md)|Returns an **Object** that has the focus. Read-only.|
| [BackColor](Outlook.Frame.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [BorderColor](Outlook.Frame.bordercolor.md)|Returns or sets a **Long** that specifies the border color of an object. Read/write.|
| [BorderStyle](Outlook.Frame.borderstyle.md)|Returns or sets an **Integer** that specifies the type of border of the control. Read/write.|
| [CanPaste](Outlook.Frame.canpaste.md)|Returns a **Boolean** that specifies whether the Clipboard contains data that the object supports. Read-only.|
| [CanRedo](Outlook.Frame.canredo.md)|Returns a **Boolean** that specifies if the most recent **Undo** can be reversed. Read-only.|
| [CanUndo](Outlook.Frame.canundo.md)|Returns a **Boolean** that specifies whether the last user action can be undone. Read-only.|
| [Caption](Outlook.Frame.caption.md)|Returns or sets a **String** that appears on an object to identify or describe it. Read/write.|
| [Cycle](Outlook.Frame.cycle.md)|Returns or sets an **Integer** that specifies whether cycling includes controls nested in a Frame. Read/write.|
| [Enabled](Outlook.Frame.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [ForeColor](Outlook.Frame.forecolor.md)|Returns or sets a **Long** that specifies the foreground color of an object. Read/write.|
| [InsideHeight](Outlook.Frame.insideheight.md)|Returns a **Long** that specifies the height, in [points](../language/glossary/vbe-glossary.md#point), of the client region inside a **Frame**. Read-only.|
| [InsideWidth](Outlook.Frame.insidewidth.md)|Returns a **Long** that specifies the width, in [points](../language/glossary/vbe-glossary.md#point), of the client region inside a **Frame**. Read-only.|
| [KeepScrollBarsVisible](Outlook.Frame.keepscrollbarsvisible.md)|Returns or sets an **Integer** that specifies whether scroll bars remain visible when not required. Read/write.|
| [MouseIcon](Outlook.Frame.mouseicon.md)|Returns a **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](Outlook.Frame.mousepointer.md)|Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [Picture](Outlook.Frame.picture.md)|Returns a **String** that specifies the full path name of a bitmap to display on a control. Read-only.|
| [PictureAlignment](Outlook.Frame.picturealignment.md)|Returns or sets an **Integer** that specifies the location of a background picture. Read/write.|
| [PictureSizeMode](Outlook.Frame.picturesizemode.md)|Returns or sets an **Integer** that specifies how to display the background picture on a **Frame**. Read/write.|
| [PictureTiling](Outlook.Frame.picturetiling.md)|Returns or sets a **Boolean** that specifies whether a picture is repeated across the background of the object. Read/write.|
| [ScrollBars](Outlook.Frame.scrollbars.md)|Returns or sets an **Integer** that specifies whether a control has vertical scroll bars, horizontal scroll bars, or both. Read/write.|
| [ScrollHeight](Outlook.Frame.scrollheight.md)|Returns or sets a **Single** that specifies the height, in [points](../language/glossary/vbe-glossary.md#point), of the total area that can be viewed by moving the scroll bars on the **Frame**. Read/write.|
| [ScrollLeft](Outlook.Frame.scrollleft.md)|Returns or sets a **Single** that specifies the distance, in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the visible form from the left edge of the **Frame**. Read/write.|
| [ScrollTop](Outlook.Frame.scrolltop.md)|Returns or sets a **Single** that specifies the distance, in [points](../language/glossary/vbe-glossary.md#point), of the top edge of the visible form from the top edge of the **Frame**. Read/write.|
| [ScrollWidth](Outlook.Frame.scrollwidth.md)|Returns or sets a **Single** that specifies the width, in [points](../language/glossary/vbe-glossary.md#point), of the total area that can be viewed by moving the scroll bars on the **Frame**. Read/write.|
| [SpecialEffect](Outlook.Frame.specialeffect.md)|Returns or sets an **Integer** that specifies the visual appearance of an object. Read/write.|
| [VerticalScrollBarSide](Outlook.Frame.verticalscrollbarside.md)|Returns or sets an **Integer** that specifies whether a vertical scroll bar appears on the right or left side of a frame. Read/write.|
| [Zoom](Outlook.Frame.zoom.md)|Returns or sets an **Integer** that specifies the percentage to increase or decrease the displayed image. Read/write.|



## Events

|Name|Description|
|:-----|:-----|
| [Click](Outlook.Frame.click.md)|Occurs when the user clicks inside the control.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]