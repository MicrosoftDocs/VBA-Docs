---
title: CommandBarButton.HyperlinkType property (Office)
keywords: vbaof11.chm6008
f1_keywords:
- vbaof11.chm6008
ms.prod: office
api_name:
- Office.CommandBarButton.HyperlinkType
ms.assetid: 5769ce22-a9e8-3eb2-919f-a3d016cf0706
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.HyperlinkType property (Office)

Sets or gets a **[msoCommandBarButtonHyperlinkType](office.msocommandbarbuttonhyperlinktype.md)** constant that represents the type of hyperlink associated with the specified command bar button. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**HyperlinkType**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Example

This example checks the **HyperlinkType** property for the specified command bar button on the command bar named **Custom**. If **HyperlinkType** is set to **msoCommandBarButtonHyperlinkNone**, the example sets the property to **msoCommandBarButtonHyperlinkOpen** and sets the URL to `www.microsoft.com`.


```vb
Set myBar = CommandBars _ 
    .Add(Name:="Custom", Position:=msoBarTop, _ 
    Temporary:=True) 
Set myButton = myBar.Controls.Add(Type:=msoControlButton) 
With myButton 
    .FaceId = 277 
    .HyperlinkType = msoCommandBarButtonHyperlinkNone 
End With 
If myButton.HyperlinkType > _ 
    msoCommandBarButtonHyperlinkOpen Then 
    myButton.HyperlinkType = _ 
        msoCommandBarButtonHyperlinkOpen 
    myButton.TooltipText = "www.microsoft.com" 
End If
```


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]