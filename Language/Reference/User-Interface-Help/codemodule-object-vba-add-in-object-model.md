---
title: CodeModule Object (VBA Add-In Object Model)
keywords: vbob6.chm1070944
f1_keywords:
- vbob6.chm1070944
ms.prod: office
ms.assetid: f2ce876d-ee2b-058f-37fc-f681bd41f139
ms.date: 06/08/2017
---


# CodeModule Object (VBA Add-In Object Model)



Represents the code behind a component, such as a [form](../../Glossary/vbe-glossary.md#form), [class](../../Glossary/vbe-glossary.md#class), or [document](../../Glossary/vbe-glossary.md#document).

## Remarks

You use the  **CodeModule** object to modify (add, delete, or edit) the code associated with a component.
Each component is associated with one  **CodeModule** object. However, a **CodeModule** object can be associated with multiple[code panes](../../Glossary/vbe-glossary.md#code-pane).
The methods associated with the  **CodeModule** object enable you to manipulate and return information about the code text on a line-by-line basis. For example, you can use the **AddFromString** method to add text to the[module](../../Glossary/vbe-glossary.md#module).  **AddFromString** places the text just above the first[procedure](../../Glossary/vbe-glossary.md#procedure) in the module or places the text at the end of the module if there are no procedures.
Use the  **Parent** property to return the **VBComponent** object associated with a[code module](../../Glossary/vbe-glossary.md#code-module).

