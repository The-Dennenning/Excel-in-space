Attribute VB_Name = "In_which_you_hear_a_clapping_sound"
Option Explicit
'In which you hear a clapping sound
'[trigger: move]

Public Sub T0()

Dim text as string

text = "A loud clap eminates from your right - do you move towards it, or away from it?"

UserForm_Button2.Label1.Caption = text
UserForm_Button2.Label2.Caption = "T0"
UserForm_Button2.CommandButton_Option1.caption = "Towards"
UserForm_Button2.CommandButton_Option2.caption = "Away"
UserForm_Button2.Show

End Sub

Public Sub T0x0()

Dim text as string

text = "You move towards the noise, though you find nothing but a small, winged automaton, clapping it's wings together as it sits delicately on a spur of rock. It looks at you suspiciously. Do you investigate it, harass it, or leave it be?"

UserForm_Button3.Label1.Caption = text
UserForm_Button3.Label2.Caption = "T0x0"
UserForm_Button3.CommandButton_Option1.caption = "Investigate"
UserForm_Button3.CommandButton_Option2.caption = "Harass"
UserForm_Button3.CommandButton_Option3.caption = "Leave"
UserForm_Button3.Show

End Sub

Public Sub T0x1()

Dim text as string

text = "You immediately walk in the opposite direction, and you feel safer. [Do: Personality, gain, neuroticism, 2]"

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x1"
UserForm_Button1.Show

End Sub

Public Sub T0x1x0()

Dim text as string

text = "Unfortunately, the clapping follows you. In a rare moment of courage, you turn to face the source of these auditory disturbances, though you only find a small, winged automaton. You are delighted by it's diminuitive size. [Do: knowledge, gain, robotics, 1]"

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x1x0"
UserForm_Button1.Show

End Sub

Public Sub T0x0x0()

Dim text as string

text = "Zooming in your optical sensors, you marvel at the fine joints of the creature. Whatever manufactured it employed immaculate attention to detail.  "
call do_action("knowledge", "gain\robotics\5")

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0x0"
UserForm_Button1.Show

End Sub

Public Sub T0x0x1()

Dim text as string

text = "You first try to put it down with words, but the automaton doesn't react. It might not have auditory sensors that would pick up the frequencies that your voicebox produces. You switch to beaming unpleasantries via your transmitter. The automaton flies away, adequately fed up with your vitriol. [Do: personality, lose, extraversion, 5]"

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0x1"
UserForm_Button1.Show

End Sub

Public Sub T0x0x2()

Dim text as string

text = "You leave it be, figuring such a small creature wouldn't take kindly to your advances. "
call do_action("personality", "gain\extraversion\2")

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0x2"
UserForm_Button1.Show

End Sub

Public Sub direct(name as string, caption as string)

If name = "T0" and caption = "Towards" then
	Call T0x0
	exit sub
End If

If name = "T0" and caption = "Away" then
	Call T0x1
	exit sub
End If

If name = "T0x1" then
	Call T0x1x0
	exit sub
End If

If name = "T0x0" and caption = "Investigate" then
	Call T0x0x0
	exit sub
End If

If name = "T0x0" and caption = "Harass" then
	Call T0x0x1
	exit sub
End If

If name = "T0x0" and caption = "Leave" then
	Call T0x0x2
	exit sub
End If

End Sub
