Attribute VB_Name = "In_which_you_learn_something"
Option Explicit
'In which you learn something
'[trigger: move]

Public Sub T0()

Dim text as string

text = "You come across an old looking tablet, it's stone edges worn with immense age. It's lexigraphic qualities are intoxicating at first glance. Do you try to read it, or toss it as far as you can?"

UserForm_Button2.Label1.Caption = text
UserForm_Button2.Label2.Caption = "T0"
UserForm_Button2.CommandButton_Option1.caption = "Read"
UserForm_Button2.CommandButton_Option2.caption = "Toss"
UserForm_Button2.Show

End Sub

Public Sub T0x0()

Dim text as string

text = "You focus your optical sensors on the old text."

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0"
UserForm_Button1.Show

End Sub

Public Sub T0x1()

Dim text as string

text = "Winding up, you toss the heavy stone a solid 4 nanolightseconds. You feel more spontaneous. [Do: Personality, lose, conscientiousness, 1]"

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x1"
UserForm_Button1.Show

End Sub

Public Sub T0x0x0()

Dim text as string

text = "Unfortunately, the scribbles etched on the stone do nothing more than tickle your processor with meaningless quantum foam. You don't get very far in attempting to decode or understand the text, though you feel good for trying. "

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0x0"
UserForm_Button1.Show

End Sub

Public Sub T0x0x1()

Dim text as string

text = "With your knowledge regarding the inner machinations of language, you manage to learn a little about the coding language behind the ancient race of robots who carved this tablet in the hopes that a future being would read it. [Do: Knowledge, gain, computer science, 5]"

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0x1"
UserForm_Button1.Show

End Sub

Public Sub direct(name as string, caption as string)

If name = "T0" and caption = "Read" then
	Call T0x0
	exit sub
End If

If name = "T0" and caption = "Toss" then
	Call T0x1
	exit sub
End If

If name = "T0x0" and KnowCheck("Language") < 50 then
	Call T0x0x0
	exit sub
End If

If name = "T0x0" and KnowCheck("Language")>=50 then
	Call T0x0x1
	exit sub
End If

End Sub
