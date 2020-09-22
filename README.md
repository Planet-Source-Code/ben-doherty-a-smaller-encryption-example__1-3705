<div align="center">

## A smaller encryption example


</div>

### Description

Takes a string and shifts each character in that string according to a password set by the user. Much smaller than most of the examples of PSC.
 
### More Info
 
you need to put in 2 text boxes (text1 -string to encrypt) (text2 - where the password goes) and 2 command buttons cmdEncrypt and cmdDecrypt then just copy and paste the following code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ben Doherty](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ben-doherty.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ben-doherty-a-smaller-encryption-example__1-3705/archive/master.zip)





### Source Code

```
Private Sub cmdEncrypt_Click()
pass$ = Len(password.Text) 'the number you shift each letter to encrypt
tmpstr = Len(Text1.Text)
If tmpstr = "0" Then
MsgBox ("You must first type in something to Encrypt") 'You can't encrypt nothing
Exit Sub
End If
For i = 1 To tmpstr
letter = Mid$(Text1.Text, i, 1)   'takes the ascii value and adds the length of the password to it
encstr = Asc(letter) + pass$
newstr = Chr$(encstr)    'changes ascii value to a character
encrypted$ = encrypted$ & newstr 'puts all the encrypted characters together
Next i
Text1.Text = encrypted$  'puts the encrypted string in text box
End Sub
Private Sub cmdDecrypt_Click()
pass$ = Len(password.Text)        'this is the exact same for the Encrypt Function
tmpstr = Len(Text1.Text)        'the only difference is that instead of adding the lenght of password.text
              'it is subtracted
If tmpstr = "0" Then
MsgBox ("You must first type in something to Decrypt")
Exit Sub
End If
For i = 1 To tmpstr
letter = Mid$(Text1.Text, i, 1)
encstr = Asc(letter) - pass$
newstr = Chr$(encstr)
decrypted$ = decrypted$ & newstr
Next i
Text1.Text = decrypted$
End Sub
```

