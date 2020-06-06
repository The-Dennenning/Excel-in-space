Attribute VB_Name = "In_which_you_nearly_kick_a_metal_rod"
Option Explicit
'In which you nearly kick a metal rod
'[must: have legs][must: have manipulator][must: have arm][trigger: move]

Public Sub T0()

Dim text as string

text = "As you walk along the " & ref("Planet", "Landscape_Adjective") & " terrain, you nearly kick what looks to be a metal rod protruding from the side of the pathway. Do you investigate further, or leave it be?"

UserForm_Button2.Label1.Caption = text
UserForm_Button2.Label2.Caption = "T0"
UserForm_Button2.CommandButton_Option1.caption = "Investigate"
UserForm_Button2.CommandButton_Option2.caption = "Leave"
UserForm_Button2.Show

End Sub

Public Sub T0x0()

Dim text as string

text = "You stoop down to get a closer look at the thing."

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0"
UserForm_Button1.Show

End Sub

Public Sub T0x1()

Dim text as string

text = "You walk away.  "
call do_action("Personality", "Lose\Openness\2")

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x1"
UserForm_Button1.Show

End Sub

Public Sub T0x0x0()

Dim text as string

text = "You notice that the rod has some wires plugged in to a small port just above where it's buried. The wires trace down the rod, probably terminating somewhere else on the instrument. It looks like some sort of electrified club. Do you cut the wires and pull it out, or walk away?"

UserForm_Button2.Label1.Caption = text
UserForm_Button2.Label2.Caption = "T0x0x0"
UserForm_Button2.CommandButton_Option1.caption = "Cut"
UserForm_Button2.CommandButton_Option2.caption = "Walk"
UserForm_Button2.Show

End Sub

Public Sub T0x0x1()

Dim text as string

text = "The rod looks pretty nondescript - it looks like an ordinary piece of metal. Do you pull it out, or walk away?"

UserForm_Button2.Label1.Caption = text
UserForm_Button2.Label2.Caption = "T0x0x1"
UserForm_Button2.CommandButton_Option1.caption = "pull"
UserForm_Button2.CommandButton_Option2.caption = "Walk"
UserForm_Button2.Show

End Sub

Public Sub T0x0x0x0()

Dim text as string

text = "You deftly slice the wires with a short-ranged laser burst from your manipulator, then tug the metal rod from the ground. It comes out cleanly. It is indeed a weapon - there's a handle where it can be grabbed, as well as a universal port for an arm to plug into if one desires a more permanent grip on the thing.  "
call do_action("gain_part", "part\SCV_ShockBaton_01\broken")

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0x0x0"
UserForm_Button1.Show

End Sub

Public Sub T0x0x1x0()

Dim text as string

text = "You grip the rod, and are immediately alerted to the electricity flowing through your arm from the rod. You pull, and it comes out cleanly, though your arm may be a little damaged.    "
call do_action("Player", "damage\arm\20%")
call do_action("Personality", "Gain\Openness\2")

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0x1x0"
UserForm_Button1.Show

End Sub

Public Sub T0x0x1x1()

Dim text as string

text = "You don't bother with it - after all, you obviously have more important things on your mind - and walk away. You feel as if you've gained nothing from this experience."

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0x1x1"
UserForm_Button1.Show

End Sub

Public Sub T0x0x1x0x0()

Dim text as string

text = "Looking at the rod more closely, you see it is a weapon - an electrified club - and as your arm can attest to, it is in good working order. There's a handle where it can be grabbed, as well as a universal port for an arm to plug into if one desires a more permanent grip on the thing.  "
call do_action("part", "gain\SCV_ShockBaton_01")

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0x1x0x0"
UserForm_Button1.Show

End Sub

Public Sub direct(name as string, caption as string)

If name = "T0" and caption = "Investigate" then
	Call T0x0
	exit sub
End If

If name = "T0" and caption = "Leave" then
	Call T0x1
	exit sub
End If

If name = "T0x0" and KnowCheck("robotics") > 25 then
	Call T0x0x0
	exit sub
End If

If name = "T0x0" and KnowCheck("robotics") <= 25 then
	Call T0x0x1
	exit sub
End If

If name = "T0x0x0" and caption = "Cut" then
	Call T0x0x0x0
	exit sub
End If

If name = "T0x0x1" and caption = "pull" then
	Call T0x0x1x0
	exit sub
End If

If name = "T0x0x1" and caption = "Walk" then
	Call T0x0x1x1
	exit sub
End If

If name = "T0x0x1x0" then
	Call T0x0x1x0x0
	exit sub
End If

If name = "T0x0x0" and caption = "Walk" then
	Call T0x0x1x1
	exit sub
End If

End Sub
