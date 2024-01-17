p=0
k=0
do
randomize
a=int(rnd()*100)
b=int(rnd()*100)
c=int(rnd()*1000)
d=int(rnd()*1000)
e=int(rnd()*4)
f=int(rnd()*4)
if e>1 then
	if e>2 then
		if e>3 then
			x=a*b
			n=a&"*"&b&"+"
		else
			x=a/b
			n=a&"/"&b&"+"
		end if
	else
		x=a-b
		n=a&"-"&b&"+"
	end if
else
	x=a+b
	n=a&"+"&b&"+"
end if
if f>1 then
	if f>2 then
		if f>3 then
			y=c*d
			m=c&"*"&d&"="
		else
			y=c+d
			m=c&"+"&d&"="
		end if
	else
		y=c/d
		m=c&"/"&d&"="
	end if
else
	y=c-d
	m=c&"-"&d&"="
end if
z=x+y
t=n+m
i=inputbox(t)
if abs(i*1000)=abs(z*1000) then
	k=k+1
	msgbox "T"
else
	msgbox "F，"&z
	if k>0 then
		k=k-1
	else
		if k>-10 then
			k=k-0.1
		else
			k=-15
		end if
	end if
end if
p=p+1
loop while k<10.8
msgbox "times:"&p&" score:"&k
