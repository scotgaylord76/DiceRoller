<html>
<head>
<title> Dice Test </title>
</head>
<body bgcolor= "white" text ="black">
<%
Dim Dice1, Dice2, Dice3, Phrase, Name1, Name2, Name3, Name4, Name5, Name6

'kkaggkd 

Randomize
Dice1 = Int(Rnd * 6) + 1
Dice2 = Int(Rnd * 6) + 1
Dice3 = Int(Rnd * 6) + 1 

Name1 = "Tim"
Name2 = "Scot"
Name3 = "Dylan"
Name4 = "Cory"
Name5 = "Tom"
Name6 = "Andrew"

%>

<% If Dice1 = 1 Then
  Set Dice1 = Name1 %>
<% ElseIf Dice1 = 2 Then
  Set Dice1 = Name2 %>
<% ElseIf Dice1 = 3 Then
  Set Dice1 = Name3 %>
<% ElseIf Dice1 = 4 Then
  Set Dice1 = Name4 %>
<% ElseIf Dice1 = 5 Then
  Set Dice1 = Name5 %>
<% ElseIf Dice1 = 6 Then
  Set Dice1 = Name6 %>
<% End If%>

<%=Dice1%>


</html>
