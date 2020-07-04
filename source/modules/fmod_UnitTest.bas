Option Compare Database
Option Explicit

Sub UnitTest1()
    Dim CC As clsMsgBox
    Dim iR As Integer
    
    Set CC = New clsMsgBox
    iR = CC.MessageBoxEx("Do you want to save the changes you made to whatever?", Exclamation + DefaultButton2, , "&Save", "Do&n't Save", "&Cancel")
    If iR = Button1 Then
        Debug.Print "Button1 Clicked"
    ElseIf iR = Button2 Then
        Debug.Print "Button2 Clicked"
    ElseIf iR = Button3 Then
        Debug.Print "Button3 Clicked"
    End If

    Set CC = New clsMsgBox
    CC.Title = "Title"
    CC.Prompt = "Prompt"
    CC.icon = Question + DefaultButton3
    CC.ButtonText1 = "ButtonText1"
    CC.ButtonText2 = "ButtonText2"
    CC.ButtonText3 = "ButtonText3"
    iR = CC.MessageBox()
    If iR = Button1 Then
        Debug.Print "Button1 Clicked"
    ElseIf iR = Button2 Then
        Debug.Print "Button2 Clicked"
    ElseIf iR = Button3 Then
        Debug.Print "Button3 Clicked"
    End If

    Set CC = New clsMsgBox
    CC.Title = "Title"
    CC.Prompt = "Prompt"
    CC.icon = Exclamation
    CC.ButtonText1 = "ButtonText1"
    iR = CC.MessageBox()
    If iR = Button1 Then
        Debug.Print "Button1 Clicked"
    ElseIf iR = Button2 Then
        Debug.Print "Button2 Clicked"
    ElseIf iR = Button3 Then
        Debug.Print "Button3 Clicked"
    End If

    Set CC = New clsMsgBox
    CC.Title = "NoIconTitle"
    CC.Prompt = "NoIconPrompt"
    CC.icon = NoIcon
    CC.ButtonText1 = "ButtonText1"
    CC.ButtonText2 = "ButtonText2"
    iR = CC.MessageBox()
    If iR = Button1 Then
        Debug.Print "Button1 Clicked"
    ElseIf iR = Button2 Then
        Debug.Print "Button2 Clicked"
    ElseIf iR = Button3 Then
        Debug.Print "Button3 Clicked"
    End If
    
    Set CC = New clsMsgBox
    CC.ButtonText1 = "No Options"
    iR = CC.MessageBox()
    If iR = Button1 Then
        Debug.Print "Button1 Clicked"
    ElseIf iR = Button2 Then
        Debug.Print "Button2 Clicked"
    ElseIf iR = Button3 Then
        Debug.Print "Button3 Clicked"
    End If
End Sub
Sub UnitTest2()
    Dim CC As clsMsgBox
    Dim iR As Integer
    
    Set CC = New clsMsgBox
        
    CC.UseCancel = True
    iR = CC.MessageBoxEx("Do you want to save the changes you made to whatever?", Exclamation + DefaultButton2, , "&Save", "Do&n't Save", "&Cancel")
    If iR = Button1 Then
        Debug.Print "Button1 Clicked"
    ElseIf iR = Button2 Then
        Debug.Print "Button2 Clicked"
    ElseIf iR = Button3 Then
        Debug.Print "Cancelled"
        Debug.Print "Button3 Clicked"
    End If

    Set CC = New clsMsgBox
    CC.UseCancel = True
    CC.Title = "Title"
    CC.Prompt = "Prompt"
    CC.icon = Question + DefaultButton3
    CC.ButtonText1 = "ButtonText1"
    CC.ButtonText2 = "ButtonText2"
    CC.ButtonText3 = "ButtonText3"
    iR = CC.MessageBox()
    If iR = Button1 Then
        Debug.Print "Button1 Clicked"
    ElseIf iR = Button2 Then
        Debug.Print "Button2 Clicked"
    ElseIf iR = Button3 Then
        Debug.Print "Cancelled"
        Debug.Print "Button3 Clicked"
    End If

    Set CC = New clsMsgBox
    CC.UseCancel = True
    CC.Title = "Title"
    CC.Prompt = "Prompt"
    CC.icon = Exclamation
    CC.ButtonText1 = "ButtonText1"
    iR = CC.MessageBox()
    If iR = Button1 Then
        Debug.Print "Button1 Clicked"
    End If

    Set CC = New clsMsgBox
    CC.UseCancel = True
    CC.Title = "NoIconTitle"
    CC.Prompt = "NoIconPrompt"
    CC.icon = NoIcon
    CC.ButtonText1 = "ButtonText1"
    CC.ButtonText2 = "ButtonText2"
    iR = CC.MessageBox()
    If iR = Button1 Then
        Debug.Print "Button1 Clicked"
    ElseIf iR = Button2 Then
        Debug.Print "Cancelled"
        Debug.Print "Button2 Clicked"
    End If
    
    Set CC = New clsMsgBox
    CC.UseCancel = True
    CC.ButtonText1 = "No Options"
    iR = CC.MessageBox()
    If iR = Button1 Then
        Debug.Print "Button1 Clicked"
    End If
End Sub

Public Function Test()
    UnitTest1
    UnitTest2
End Function