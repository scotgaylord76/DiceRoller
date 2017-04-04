<%
Dim adoCon
Dim rsAddComments
Dim strSQL
Dim Die1, Die2, Die3
Dim Name1, Name2, Name3, Name4, Name5, Name6
Dim Verb1, Verb2, Verb3, Verb4, Verb5, Verb6
Dim Adj1, Adj2, Adj3, Adj4, Adj5, Adj6
Dim NameRoll
Dim VerbRoll
Dim AdjRoll
Dim Phrase
Randomize

Name1 = "Tim "
Name2 = "Scot "
Name3 = "Andrew "
Name4 = "Tom "
Name5 = "Dylan "
Name6 = "Cory "

Verb1 = "is "
Verb2 = "was "
Verb3 = "smells "
Verb4 = "tastes "
Verb5 = "looks "
Verb6 = "acts "

Adj1 = "awesome!"
Adj2 = "great!"
Adj3 = "bad..."
Adj4 = "terrible!"
Adj5 = "okay..."
Adj6 = "mediocre..."

Die1 = Int(Rnd * 6) + 1
Die2 = Int(Rnd * 6) + 1
Die3 = Int(Rnd * 6) + 1

If Die1 = 1 Then 
 NameRoll = Name1
ElseIf Die1 = 2 Then
 NameRoll = Name2
ElseIf Die1 = 3 Then
 NameRoll = Name3
ElseIf Die1 = 4 Then
 NameRoll = Name4
ElseIf Die1 = 5 Then
 NameRoll = Name5
Else NameRoll = Name6
End If

If Die2 = 1 Then
 VerbRoll = Verb1
ElseIf Die2 = 2 Then
 VerbRoll = Verb2
ElseIf Die2 = 3 Then
 VerbRoll = Verb3
ElseIf Die2 = 4 Then
 VerbRoll = Verb4
ElseIf Die2 = 5 Then
 VerbRoll = Verb5
Else VerbRoll = Verb6
End If

If Die3 = 1 Then
 AdjRoll = Adj1
ElseIf Die3 = 2 Then
 AdjRoll = Adj2
ElseIf Die3 = 3 Then
 AdjRoll = Adj3
ElseIf Die3 = 4 Then
 AdjRoll = Adj4
ElseIf Die3 = 5 Then
 AdjRoll = Adj5
Else AdjRoll = Adj6
End If

Phrase = NameRoll & VerbRoll & AdjRoll

Set adoCon=Server.CreateObject("ADODB.Connection")
adoCon.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("TestDB.mdb")
Set rsAddComments = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT tblDiceLib.Dice1, tblDiceLib.Dice2, tblDiceLib.Dice3, tblDiceLib.Phrase, tblDiceLib.NameRoll, tblDiceLib.VerbRoll, tblDiceLib.AdjRoll FROM tblDiceLib;"
rsAddComments.CursorType = 2
rsAddComments.LockType = 3
rsAddComments.Open strSQL, adoCon
rsAddComments.AddNew

rsAddComments.Fields("Dice1") = Die1
rsAddComments.Fields("Dice2") = Die2
rsAddComments.Fields("Dice3") = Die3
rsAddComments.Fields("NameRoll") = NameRoll
rsAddComments.Fields("VerbRoll") = VerbRoll
rsAddComments.Fields("AdjRoll") = AdjRoll
rsAddComments.Fields("Phrase") = Phrase

rsAddComments.Update

rsAddComments.Close
Set rsAddComments = Nothing
Set adoCon = Nothing

Response.Redirect "DiceLib.asp"
%>