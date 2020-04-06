---
title: AutoCaption object (Word)
keywords: vbawd10.chm2427
f1_keywords:
- vbawd10.chm2427
ms.prod: word
api_name:
- Word.AutoCaption
ms.assetid: 895b5181-d36f-7f63-572a-c2d37c878e17
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCaption object (Word)

Represents a single caption that can be automatically added when items such as tables, pictures, or OLE objects are inserted into a document. The  **AutoCaption** object is a member of the **[AutoCaptions](Word.autocaptions.md)** collection. The **AutoCaptions** collection contains all the captions listed in the **AutoCaption** dialog box.


## Remarks

Use  **[AutoCaptions](Word.Application.AutoCaptions.md)** (_index_), where _index_ is the caption name or index number, to return a single **AutoCaption** object. The caption names correspond to the items listed in the **AutoCaption** dialog box. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown in the **AutoCaption** dialog box. The following example enables autocaptions for Word tables.


```vb
AutoCaptions("Microsoft Word Table").AutoInsert = True
```

The index number represents the position of the  **AutoCaption** object in the list of items in the **AutoCaption** dialog box. The following example displays the name of the first item listed in the **AutoCaption** dialog box.




```vb
MsgBox AutoCaptions(1).Name
```

 **AutoCaption** objects cannot be programmatically added to or deleted from the **AutoCaptions** collection.

## Properties

- [Application](Word.AutoCaption.Application.md)
- [AutoInsert](Word.AutoCaption.AutoInsert.md)
- [CaptionLabel](Word.AutoCaption.CaptionLabel.md)
- [Creator](Word.AutoCaption.Creator.md)
- [Index](Word.AutoCaption.Index.md)
- [Name](Word.AutoCaption.Name.md)
- [Parent](Word.AutoCaption.Parent.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]