msgbox "Please make sure they are in the same length",,"Attention"
inpt1=inputbox("1")
inpt2=inputbox("2")
'=======
Function CutAndJoin(sSource, iLong, sJoiner)
Dim I, N
N = Len(sSource) / iLong
If(N <> Fix(N))Then N = Fix(N) + 1
For I = 0 To N - 1
CutAndJoin = CutAndJoin & Mid(sSource, I * iLong + 1, iLong) & sJoiner
Next
If(N > 0)Then CutAndJoin = Left(CutAndJoin, Len(CutAndJoin) - Len(sJoiner))
End Function
'=============
wtxt1=CutAndJoin(inpt1, 1, "$$")
arry1=split(wtxt1,"$$")
wtxt2=CutAndJoin(inpt2, 1, "$$")
arry2=split(wtxt2,"$$")
lon1=UBound(arry1)
lon2=UBound(arry2)
'===============
simlr=0
for i=LBound(arry1) to UBound(arry1)
cmpr1=arry1(i)
cmpr2=arry2(i)
if cmpr1 = cmpr2 then
simlr=1+simlr
end if
next
smlr=int(1000000*simlr/(UBound(arry1)-LBound(arry2)+1))/10000
msgbox "Similarity="&smlr&"%"
