Attribute VB_Name = "In_which_a_drone_lands_on_your_chasis"
Option Explicit
'In which a drone lands on your chasis
'[trigger: move]
Public globalRand(1 to 32) as integer

Public Sub T0()

Dim text as string

globalRand(1) = worksheetfunction.randbetween(1, 10)

text = "A small drone lands on your chasis and begins draining electrical energy from your battery via your exposed electrical uplink. You balk at the lack of physical security of your own body. It drains " & globalRand(1) & " kilojoules. "
call do_action("energy", "lose\" & globalRand(1) & "")

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0"
UserForm_Button1.Show

End Sub

Public Sub T0x0()

Dim text as string


text = "The drone, having finished it's meal of your battery's electricity, heads on it's merry way."

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0"
UserForm_Button1.Show

End Sub

Public Sub T0x1()

Dim text as string


text = "The feeling of energy leaving your system is abstractly discomforting, though it nonetheless feels a little satisfying. You shoo the drone, and the feeling stops."

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x1"
UserForm_Button1.Show

End Sub

Public Sub direct(name as string, caption as string)

If name = "T0" and globalRand(1) < 6 then
	Call T0x0
	exit sub
End If

If name = "T0" and globalRand(1) > 5 then
	Call T0x1
	exit sub
End If

End Sub
