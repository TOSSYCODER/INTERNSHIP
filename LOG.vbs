Dim USERS
Dim USERNAME
Dim PASSWORD
Dim strMessage1
Dim strMessage2 
Dim msg
Dim objShell
Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set f = fso.CreateTextFile("C:\Users\TOUSIF\Desktop\vbs\USERS.txt", 1)


Set objShell = CreateObject("WScript.Shell")

USERNAME="admin"
PASSWORD="1234"
uincorrect = "USERNAME INCORRECT!"
pincorrect = "PASSWORD INCORRECT!"

Do
strMessage1 =Inputbox("Enter USERNAME","Input Required")
if strMessage1=USERNAME Then
USERS="USER TRIED TO LOGIN"
f.WriteLine USERS
elseif IsEmpty(strMessage1) Then
        WScript.Quit
else
msgbox(uincorrect)
USERS="INCORRECT USERNAME ENTERED!"
f.WriteLine USERS
USERS="ENTERED USERNAME: "&strMessage1
f.WriteLine USERS
end if
Loop Until strMessage1=USERNAME

if IsEmpty(strMessage1) Then
WScript.Quit
else
Do
strMessage2 =Inputbox("Enter PASSWORD","Input Required")
if strMessage2=PASSWORD Then
elseif IsEmpty(strMessage2) Then
        WScript.Quit
msgbox(pincorrect)
else
msgbox(pincorrect)
USERS="INCORRECT PASSWORD ENTERED!"
f.WriteLine USERS
USERS="ENTERED PASSWORD: "&strMessage2
f.WriteLine USERS
end if
Loop Until strMessage2=PASSWORD
end if

if IsEmpty(strMessage2) Then
WScript.Quit
else
msg="LOGIN SUCCESSFULL!"
variable=msgbox(msg)
USERS="USER LOGGED-IN SUCCESSFULL"
f.WriteLine USERS
end if

f.close