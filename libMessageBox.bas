Attribute VB_Name = "libMessageBox"

' //////////////////////////////////////////////////////////////////////
' //
' // MessageBox Version 2.0 - new state of art
' // Copyright © 2000 Andreas Schwarz, FutureProjects Development
' // e-mail: andi@futureprojects.de , http://www.futureprojects.de
' //
' // all rights reserved.
' //
' // use instead of MsgBox : MessageBox !


Public Function MessageBox(Prompt As String, Optional Flags As MsgType = 0, Optional PromptCaption As String = "", Optional iForm As Form) As MsgResult
    MessageBox = Msg.dwMessageBox(Prompt, Flags, PromptCaption, iForm)
End Function
