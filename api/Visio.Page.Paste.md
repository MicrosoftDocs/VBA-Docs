---
title: Page.Paste Method (Visio)
keywords: vis_sdr.chm10916430
f1_keywords:
- vis_sdr.chm10916430
ms.prod: visio
api_name:
- Visio.Page.Paste
ms.assetid: 73dd3b44-1288-26d1-4956-93f187d71886
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.Paste Method (Visio)

Pastes the contents of the Clipboard into an object.


## Syntax

 _expression_. `Paste`( `_Flags_` )

 _expression_ A variable that represents a [Page](./Visio.Page.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Flags_|Optional| **Variant**|Determines how shapes are translated during the paste operation.|

## Return value

Nothing


## Remarks

The  **Paste** method works only with **Shape** objects that are group shapes. Use the **Type** property of a shape to determine whether it is a group.

Possible values for  _Flags_ are declared by the Visio type library in **VisCutCopyPasteCodes** , and are described in the following table.



|**Flag**|Value|Description|
|:-----|:-----|:-----|
| **visCopyPasteNormal**|&H0|Follow default copying behavior.|
| **visCopyPasteNoTranslate**|&H1|Copy shapes to their original coordinate locations.|
| **visCopyPasteCenter**|&H2|Copy shapes to the center of the page.|
| **visCopyPasteNoHealConnectors**|&H4|Do not clean up connectors attached to cut shapes.|
| **visCopyPasteNoContainerMembers**|&H8|Do not cut and copy unselected members of containers or lists.|
| **visCopyPasteNoAssociatedCallouts**|&H16|Do not cut and copy unselected callouts associated with shapes.|
| **visCopyPasteDontAddToContainers**|&H32|Do not add pasted shapes to any underlying containers.|
| **visCopyPasteNoCascade**|&H64|Do not offset shapes on copy.|

Setting  _Flags_ to **visCopyPasteNormal** is the equivalent of the behavior in the user interface. You should use **visCopyPasteNormal** and the other flags consistently. For example, if you use the value **visCopyPasteNoTranslate** to copy, you should also use that value to paste, because that is the only way to ensure that shapes are pasted to their original coordinate location.

If you need to control the format of the pasted information and (optionally) establish a link to a source file (for example, a Microsoft Word document), use the  **PasteSpecial** method.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Paste** method. It draws a rectangle, copies it, and then pastes the copy onto the drawing page.


```vb
 
Public Sub Paste_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 'Draw a rectangle. 
 Set vsoShape = ActivePage.DrawRectangle(1, 5, 5, 1) 
 
 'Copy the shape to the Clipboard. 
 vsoShape.Copy 
 
 'Paste the copy onto the drawing page. 
 ActivePage.Paste 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]