<html>
<head>
<title>Dice of Truths</title>
</head>
<body bgcolor="white" text="black">

<%
Dim adoCon
Dim rsTestDB
Dim strSQL
'I added a comment
Set adoCon = Server.CreateObject("ADODB.Connection")
adoCon.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("TestDB.mdb")
Set rsTestDB = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT tblDiceLib.Dice1, tblDiceLib.Dice2, tblDiceLib.Dice3, tblDiceLib.Phrase, tblDiceLib.NameRoll, tblDiceLib.VerbRoll, tblDiceLib.AdjRoll FROM tblDiceLib;"
rsTestDB.Open strSQL, adoCon

Do While not rsTestDB.EOF
 Response.Write ("<br>")
 Response.Write (rsTestDB("Dice1"))
 Response.Write ("<br>")
 Response.Write (rsTestDB("Dice2"))
 Response.Write ("<br>")
 Response.Write (rsTestDB("Dice3"))
 Response.Write ("<br>")
 Response.Write (rsTestDB("NameRoll"))
 Response.Write ("<br>")
 Response.Write (rsTestDB("VerbRoll"))
 Response.Write ("<br>")
 Response.Write (rsTestDB("AdjRoll"))
 Response.Write ("<br>")
 Response.Write (rsTestDB("Phrase"))
 Response.Write ("<br>")

 rsTestDB.MoveNext
Loop

rsTestDB.Close
Set rsTestDB = Nothing
Set adoCon = Nothing
%>

</body>
</html>