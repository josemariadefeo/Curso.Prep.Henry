option explicit
Public Function sumnumer()
dim numer, a, b, c, d as int
numer =msgbox ("Ingrese un numero de 4 cifras")
a=numer mod 10
numer= int(numer/10)
b=numer mod 10
numer=int(numer/10)
c=numer mod 10
numer= int(numer/10)
d=numer
msgbox("u de mil=" & d & vbcrlf & "centena= " & c & vbcrlf & "decena= "& b & vbcrlf & "unidad= " & a)
    
End Function