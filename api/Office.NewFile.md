---
title: NewFile object (Office)
keywords: vbaof11.chm235000
f1_keywords:
- vbaof11.chm235000
ms.prod: office
api_name:
- Office.NewFile
ms.assetid: 6f53ced5-4488-b67f-ca1f-729aeb790eb1
ms.date: 01/22/2019
localization_priority: Normal
---


# NewFile object (Office)

Represents items listed on the **New Item** task pane available in several Microsoft Office applications.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Remarks

You can use the **Application** property or the **Creator** property to access the **NewFile** object in each of the applications.


## Example

Use the **Add** method to add a new item to the **New Item** task pane. The following example adds an item to Word's **New Document** task pane.

```vb
Sub AddNewDocToTaskPane() 
    Application.NewDocument.Add FileName:="C:\NewDocument.doc", _ 
        Section:=msoNew, DisplayName:="New Document" 
    CommandBars("Task Pane").Visible = True  
End Sub
```

<br/>

Use the **Remove** method to remove an item from the **New Item** task pane. The following example removes the document added in the above example from Word's **New Document** task pane.

```vb
Sub RemoveDocFromTaskPane() 
    Application.NewDocument.Remove FileName:="C:\NewDocument.doc", _ 
        Section:=msoNew, DisplayName:="New Document" 
    CommandBars("Task Pane").Visible = True  
End Sub
```

> [!NOTE] 
> These examples are for Word, but you can change the **NewDocument** property for any of the properties listed and use the code in the corresponding application.


## See also

- [NewFile object members](overview/library-reference/newfile-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]