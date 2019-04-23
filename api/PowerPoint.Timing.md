---
title: Timing object (PowerPoint)
keywords: vbapp10.chm653000
f1_keywords:
- vbapp10.chm653000
ms.prod: powerpoint
api_name:
- PowerPoint.Timing
ms.assetid: 11f7dab2-f9ed-1883-ab74-93f1be481af6
ms.date: 06/08/2017
localization_priority: Normal
---


# Timing object (PowerPoint)

Represents timing properties for an animation effect.


## Remarks

Use the following read/write properties of the  **Timing** object to manipulate animation timing effects.



|**Use this property**|**To change this...**|
|:-----|:-----|
|[Accelerate](PowerPoint.Timing.Accelerate.md)|Percentage of the duration over which acceleration should take place|
|[AutoReverse](PowerPoint.Timing.AutoReverse.md)|Whether an effect should play forward and then reverse, thereby doubling the duration|
|[Decelerate](PowerPoint.Timing.Decelerate.md)|Percentage of the duration over which acceleration should take place|
|[Duration](PowerPoint.SlideShowTransition.Duration.md)|Length of animation (in seconds)|
|[RepeatCount](PowerPoint.Timing.RepeatCount.md)|Number of times to repeat the animation|
|[RepeatDuration](PowerPoint.Timing.RepeatDuration.md)|How long should the repeats last (in seconds)|
|[Restart](PowerPoint.Timing.Restart.md)|Restart behavior of an animation node|
|[RewindAtEnd](PowerPoint.Timing.RewindAtEnd.md)|Whether an objects return to its beginning position after an effect has ended|
|[SmoothStart](PowerPoint.Timing.SmoothStart.md)|Whether an effect accelerates when it starts|
|[SmoothEnd](PowerPoint.Timing.SmoothEnd.md)|Whether an effect decelerates when it ends|
|[TriggerDelayTime](PowerPoint.Timing.TriggerDelayTime.md)|Delay time from when the trigger is enabled (in seconds)|
|[TriggerShape](PowerPoint.Timing.TriggerShape.md)|Which shape is associated with the timing effect|
|[TriggerType](PowerPoint.Timing.TriggerType.md)|How the timing effect is triggered|

## Example

To return a  **Timing** object, use the [Timing](PowerPoint.AnimationBehavior.Timing.md)property of the  **[AnimationBehavior](PowerPoint.AnimationBehavior.md)** or **[Effect](PowerPoint.Effect.md)** object. The following example sets timing duration information for the main animation.


```vb
ActiveWindow.Selection.SlideRange(1).TimeLine.MainSequence(1).Timing.Duration = 5
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]