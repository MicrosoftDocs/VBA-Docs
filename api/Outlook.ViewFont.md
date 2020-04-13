---
title: ViewFont object (Outlook)
keywords: vbaol11.chm3188
f1_keywords:
- vbaol11.chm3188
ms.prod: outlook
api_name:
- Outlook.ViewFont
ms.assetid: cbd7c6ce-f49a-1627-0ad9-a019911fb47b
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewFont object (Outlook)

Represents the font used when formatting text in various portions of a view.


## Remarks

The **ViewFont** object is used by the following objects to represent font formatting information applied to the text in various portions of a view:


- The **[HeadingsFont](Outlook.businessCardView.HeadingsFont.md)** property of the **[BusinessCardView](Outlook.businessCardView.md)** object.
    
- The **[DayWeekFont](overview/Outlook.md)**, **[DayWeekTimeFont](overview/Outlook.md)**, and **[MonthFont](overview/Outlook.md)** properties of the **[CalendarView](Outlook.CalendarView.md)** object.
    
- The **[BodyFont](Outlook.CardView.BodyFont.md)** and **[HeadingsFont](Outlook.CardView.HeadingsFont.md)** properties of the **[CardView](Outlook.CardView.md)** object.
    
- The **[AutoPreviewFont](Outlook.TableView.AutoPreviewFont.md)**, **[ColumnFont](Outlook.TableView.ColumnFont.md)**, and **[RowFont](Outlook.TableView.RowFont.md)** properties of the **[TableView](Outlook.TableView.md)** object.
    
- The **[ItemFont](Outlook.TimelineView.ItemFont.md)**, **[LowerScaleFont](Outlook.TimelineView.LowerScaleFont.md)**, and **[UpperScaleFont](Outlook.TimelineView.UpperScaleFont.md)** properties of the **[TimelineView](Outlook.TimelineView.md)** object.
    

## Example

The following Visual Basic for Applications (VBA) sample increments the value of the  **[Size](Outlook.ViewFont.Size.md)** property for the **ViewFont** object returned from the **ItemFont** property for the current **TimelineView** object.


```vb
Private Sub IncreaseItemFontSize() 
 
 Dim objTimelineView As TimelineView 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTimelineView Then 
 
 
 
 ' Obtain a TimelineView object reference for the 
 
 ' current timeline view. 
 
 Set objTimelineView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Increment the Size property of the 
 
 ' ViewFont object obtained from the 
 
 ' ItemFont property, but only 
 
 ' if the font is less than 24 points 
 
 ' in size. 
 
 If objTimelineView.ItemFont.Size < 24 Then 
 
 objTimelineView.ItemFont.Size = _ 
 
 objTimelineView.ItemFont.Size + 1 
 
 
 
 ' Save the timeline view. 
 
 objTimelineView.Save 
 
 End If 
 
 End If 
 
End Sub 
 

```


## Properties



|Name|
|:-----|
|[Application](Outlook.ViewFont.Application.md)|
|[Bold](Outlook.ViewFont.Bold.md)|
|[Class](Outlook.ViewFont.Class.md)|
|[Color](Outlook.ViewFont.Color.md)|
|[ExtendedColor](Outlook.ViewFont.ExtendedColor.md)|
|[Italic](Outlook.ViewFont.Italic.md)|
|[Name](Outlook.ViewFont.Name.md)|
|[Parent](Outlook.ViewFont.Parent.md)|
|[Session](Outlook.ViewFont.Session.md)|
|[Size](Outlook.ViewFont.Size.md)|
|[Strikethrough](Outlook.ViewFont.Strikethrough.md)|
|[Underline](Outlook.ViewFont.Underline.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]