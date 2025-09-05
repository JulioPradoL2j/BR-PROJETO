' === LicenseInit Runner ===
Dim path
Set shell = WScript.CreateObject("WScript.Shell")

' Obtém JAVA_HOME
path = shell.Environment.Item("JAVA_HOME")
If path = "" Then
    MsgBox "Could not find JAVA_HOME environment variable!", vbOKOnly, "License"
Else
    If InStr(path, "\bin") = 0 Then
        path = path + "\bin\"
    Else
        path = path + "\"
    End If
    path = Replace(path, "\\", "\")
    path = Replace(path, "Program Files", "Progra~1")
End If

' === Comando principal ===
Dim command
' libs/* para pegar todos os .jar dentro de libs/
command = path & "java -Xmx512m -cp libs/*; ext.mods.security.LicenseInit"

' === Loop de execução ===
Dim exitcode
exitcode = 0
Do
    ' Executa o comando e mantém o console fechado (0 = oculto, 1 = mostra)
    exitcode = shell.Run("cmd /c " & command & " & exit", 0, True)

    ' Trata o código de saída
    If exitcode = 2 Then
        ' Reinicia
        exitcode = 2
    ElseIf exitcode <> 0 Then
        ' Qualquer outro erro encerra o loop
        exitcode = 0
        Exit Do
    End If
Loop While exitcode = 2
