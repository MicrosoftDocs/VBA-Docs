---
title: InvisibleApp.TextChanged event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.TextChanged
ms.assetid: 7212fc84-0573-22ab-3244-b0258a24d7ad
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.TextChanged event (Visio)

Occurs after the text of a shape is changed in a document.


## Syntax

_expression_.**TextChanged** (_Shape_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape whose text changed.|

## Remarks

The  **TextChanged** event is fired when the raw text of a shape changes, such as when the characters Microsoft Visio stores for the shape change. If a shape's characters change because a user is typing, the **TextChanged** event does not fire until the text editing session terminates.

When a field is added to or removed from a shape's text, its raw text changes; hence, a  **TextChanged** event fires. However, no **TextChanged** event fires when the text in a field changes. For example, a shape has a text field that shows its width. A **TextChanged** event does not fire when the shape's width changes, because the raw text stored for the shape has not changed, even though the apparent (expanded) text of the shape does change. Use the **CellChanged** event for one of the cells in the Text Fields section to detect when the text in a text field changes.

To access a shape's raw text, use the  **Text** property. To access the text of a shape in which text fields have been expanded, use the **Characters.Text** property. You can determine the location and properties of text fields in a shape's text by using the **Shape.Characters** object.

In Visio 5.0 and earlier versions, the raw characters reported by the  **Text** property for a field included four characters, the first being the escape character. Starting with Visio 2000, only a single escape character is present in the raw text stream.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).




> [!NOTE] 
> You can use VBA  **WithEvents** variables to sink the **TextChanged** event.

For performance considerations, the  **Document** object's event set does not include the **TextChanged** event. To sink the **TextChanged** event from a **Document** object (and from the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** object in a VBA project), you must use the **AddAdvise** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]