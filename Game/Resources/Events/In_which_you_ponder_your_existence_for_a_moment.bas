Attribute VB_Name = "In_which_you_ponder_your_existence_for_a_moment"
Option Explicit
'In which you ponder your existence for a moment
'[trigger: move]

Public Sub T0()

Dim text as string

text = "Suddenly a subroutine you didn't know you had engages and sends you into an existential vortex of recurrent function calls with the ultimate goal of drawing conclusions regarding the nature of the universe, and your place in it. Do you kill this rogue process immediately, or see where it goes?"

UserForm_Button2.Label1.Caption = text
UserForm_Button2.Label2.Caption = "T0"
UserForm_Button2.CommandButton_Option1.caption = "Kill"
UserForm_Button2.CommandButton_Option2.caption = "See"
UserForm_Button2.Show

End Sub

Public Sub T0x0()

Dim text as string

text = "You're already comfortable with your place in the universe, and don't feel like putting extra processing time into ruminating on such matters. You kill the process.    "
call do_action("personality", "lose\Openness\2")
call do_action("personality", "lose\Neuroticism\2")

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x0"
UserForm_Button1.Show

End Sub

Public Sub T0x1()

Dim text as string

text = "Giving into the natural flow of this process, you contemplate who you are and where you came from. Woah.  
  "
call do_action("personality", "gain\Openness\2")
call do_action("personality", "gain\Neuroticism\2")

UserForm_Button1.Label1.Caption = text
UserForm_Button1.Label2.Caption = "T0x1"
UserForm_Button1.Show

End Sub

Public Sub direct(name as string, caption as string)

If name = "T0" and caption = "Kill" then
	Call T0x0
	exit sub
End If

If name = "T0" and caption = "See" then
	Call T0x1
	exit sub
End If

End Sub
