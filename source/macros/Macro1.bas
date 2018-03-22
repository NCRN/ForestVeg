Version =196611
ColumnsShown =0
Begin
    Action ="OnError"
    Argument ="0"
End
Begin
    Action ="GoToControl"
    Argument ="=[Screen].[PreviousControl].[Name]"
End
Begin
    Action ="ClearMacroError"
End
Begin
    Condition ="Not [Form].[NewRecord]"
    Action ="RunCommand"
    Argument ="223"
End
Begin
    Condition ="[Form].[NewRecord] And Not [Form].[Dirty]"
    Action ="Beep"
End
Begin
    Condition ="[Form].[NewRecord] And [Form].[Dirty]"
    Action ="RunCommand"
    Argument ="292"
End
Begin
    Condition ="[MacroError]<>0"
    Action ="MsgBox"
    Argument ="=[MacroError].[Description]"
    Argument ="-1"
    Argument ="0"
End
Begin
    Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
        "nterfaceMacro MinimumClientDesignVersion=\"14.0.0000.0000\" xmlns=\"http://schem"
        "as.microsoft.com/office/accessservices/2009/11/application\" xmlns:a=\"http://sc"
        "hemas.microsoft.com/office/acc"
End
Begin
    Comment ="_AXL:essservices/2009/11/forms\"><Statements><Action Name=\"OnError\"/><Action N"
        "ame=\"GoToControl\"><Argument Name=\"ControlName\">=[Screen].[PreviousControl].["
        "Name]</Argument></Action><Action Name=\"ClearMacroError\"/><ConditionalBlock><If"
        "><Condition>Not [Form]"
End
Begin
    Comment ="_AXL:.[NewRecord]</Condition><Statements><Action Name=\"DeleteRecord\"/></Statem"
        "ents></If></ConditionalBlock><ConditionalBlock><If><Condition>[Form].[NewRecord]"
        " And Not [Form].[Dirty]</Condition><Statements><Action Name=\"Beep\"/></Statemen"
        "ts></If></Conditi"
End
Begin
    Comment ="_AXL:onalBlock><ConditionalBlock><If><Condition>[Form].[NewRecord] And [Form].[D"
        "irty]</Condition><Statements><Action Name=\"UndoRecord\"/></Statements></If></Co"
        "nditionalBlock><ConditionalBlock><If><Condition>[MacroError]&lt;&gt;0</Condition"
        "><Statements><A"
End
Begin
    Comment ="_AXL:ction Name=\"MessageBox\"><Argument Name=\"Message\">=[MacroError].[Descrip"
        "tion]</Argument></Action></Statements></If></ConditionalBlock></Statements></Use"
        "rInterfaceMacro>"
End
