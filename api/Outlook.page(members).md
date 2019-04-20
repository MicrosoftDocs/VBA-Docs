---
title: Page Members (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 430d453a-6727-4dc6-bc77-0ec9866b4a38
ms.date: 06/08/2017
localization_priority: Normal
---


# Page Members (Outlook Forms Script)

Represents one page of a  [MultiPage](Outlook.multipage.md) or a single member of a [Pages](Outlook.pages.md) collection.


## Methods



|Name|Description|
|:-----|:-----|
| [Copy](Outlook.Page.copy.md)|Copies the contents of an object to the Clipboard.|
| [Cut](Outlook.Page.cut.md)|Removes selected information from an object and transfers it to the Clipboard.|
| [Paste](Outlook.Page.paste.md)|Transfers the contents of the Clipboard to an object.|
| [RedoAction](Outlook.Page.redoaction.md)|Reverses the effect of the most recent  **Undo** action.|
| [Repaint](Outlook.Page.repaint.md)|Updates the display by redrawing the page.|
| [Scroll](Outlook.Page.scroll.md)|Moves the scroll bar on an object.|
| [SetDefaultTabOrder](Outlook.Page.setdefaulttaborder.md)|Sets the  **TabIndex** property of each control on a frame or page, using a default top-to-bottom, left-to-right tab order.|
| [UndoAction](Outlook.Page.undoaction.md)|Reverses the most recent action that supports the  **Undo** command.|



## Properties



|Name|Description|
|:-----|:-----|
| [Accelerator](Outlook.Page.accelerator.md)|Returns or sets the accelerator key for the page. Read/write.|
| [ActiveControl](Outlook.Page.activecontrol.md)|Returns an  **Object** that has the focus. Read-only.|
| [CanPaste](Outlook.Page.canpaste.md)|Returns a  **Boolean** that specifies whether the Clipboard contains data that the object supports. Read-only.|
| [CanRedo](Outlook.Page.canredo.md)|Returns a  **Boolean** that specifies if the most recent **Undo** can be reversed. Read-only.|
| [CanUndo](Outlook.Page.canundo.md)|Returns a  **Boolean** that specifies whether the last user action can be undone. Read-only.|
| [Caption](Outlook.Page.caption.md)|Returns or sets a  **String** that specifies the text that appears on the page. Read/write.|
| [ControlTipText](Outlook.Page.controltiptext.md)|Returns and sets a  **String** that specifies text that appears when the user briefly holds the mouse pointer over a control without clicking. Read/write.|
| [Cycle](Outlook.Page.cycle.md)|Returns or sets an  **Integer** that specifies whether cycling includes controls nested in a [MultiPage](Outlook.multipage.md). Read/write.|
| [Enabled](Outlook.Page.enabled.md)|Returns or sets a  **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [Index](Outlook.Page.index.md)|Returns or sets an  **Integer** that specifies the position of a [Page](Outlook.Page.md) object in a [Pages](Outlook.pages.md) collection. Read/write.|
| [InsideHeight](Outlook.Page.insideheight.md)|Returns a  **Long** that specifies the height, in [points](../language/glossary/vbe-glossary.md#point), of the client region inside a **Page**. Read-only.|
| [InsideWidth](Outlook.Page.insidewidth.md)|Returns a  **Long** that specifies the width, in [points](../language/glossary/vbe-glossary.md#point), of the client region inside a **Page**. Read-only.|
| [KeepScrollBarsVisible](Outlook.Page.keepscrollbarsvisible.md)|Returns or sets an  **Integer** that specifies whether scroll bars remain visible when not required. Read/write.|
| [Name](Outlook.Page.name.md)|Returns or sets a  **String** that specifies the name of an object. Read/write.|
| [Parent](Outlook.Page.parent.md)|Returns an  **Object** that represents the parent object of the specified page. Read-only.|
| [Picture](Outlook.Page.picture.md)|Returns a  **String** that specifies the full path name of a bitmap to display on a control. Read-only.|
| [PictureAlignment](Outlook.Page.picturealignment.md)|Returns or sets an  **Integer** that specifies the location of a background picture. Read/write.|
| [PictureSizeMode](Outlook.Page.picturesizemode.md)|Returns or sets an  **Integer** that specifies how to display the background picture on a page. Read/write.|
| [PictureTiling](Outlook.Page.picturetiling.md)|Returns or sets a  **Boolean** that specifies whether a picture is repeated across the background of the object. Read/write.|
| [ScrollBars](Outlook.Page.scrollbars.md)|Returns or sets an  **Integer** that specifies whether a page has vertical scroll bars, horizontal scroll bars, or both. Read/write.|
| [ScrollHeight](Outlook.Page.scrollheight.md)|Returns or sets a  **Single** that specifies the height, in [points](../language/glossary/vbe-glossary.md#point), of the total area that can be viewed by moving the scroll bars on the **Page**. Read/write.|
| [ScrollLeft](Outlook.Page.scrollleft.md)|Returns or sets a  **Single** that specifies the distance, in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the visible form from the left edge of the **Page**. Read/write.|
| [ScrollTop](Outlook.Page.scrolltop.md)|Returns or sets a  **Single** that specifies the distance, in [points](../language/glossary/vbe-glossary.md#point), of the top edge of the visible form from the top edge of the **Page**. Read/write.|
| [ScrollWidth](Outlook.Page.scrollwidth.md)|Returns or sets a  **Single** that specifies the width, in [points](../language/glossary/vbe-glossary.md#point), of the total area that can be viewed by moving the scroll bars on the **Page**. Read/write.|
| [Tag](Outlook.Page.tag.md)|Returns or sets a  **String** that specifies additional information about an object. Read/write.|
| [VerticalScrollBarSide](Outlook.Page.verticalscrollbarside.md)|Returns or sets an  **Integer** that specifies whether a vertical scroll bar appears on the right or left side of a page. Read/write.|
| [Visible](Outlook.Page.visible.md)|Returns or sets a  **Boolean** that specifies whether a **Page** is visible or hidden. Read/write.|
| [Zoom](Outlook.Page.zoom.md)|Returns or sets an  **Integer** that specifies the percentage to increase or decrease the displayed image. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]