Dim fso, objShell, objSource, item, zip_contnts

Dim i, j, k, l

i = 0

j = 0

k = 0

l = 1

'text = "a) Star office "

Set fso = WScript.CreateObject("Scripting.Filesystemobject")

Set objShell = CreateObject("Shell.Application")

pathToFile = "C:\Users\Sujatha\Desktop\Raj\CSVolume1-2\CS-Volume2.txt"

If fso.FileExists (pathToFile) = True Then

    Call Read_Data_From_Text_File(pathToFile)
    
Else 

    MsgBox "File does not exist Dummy !!!"
    
End If

Function Read_Data_From_Text_File(filepath)

    Set open_file = fso.OpenTextFile (filepath, 1)

    Do Until open_file.AtEndOfStream
    
        Data_text = open_file.ReadLine
        
        Call Condition_data(Data_text)
    
    Loop
        
    open_file.Close
    
    Call create_conditioned_text_file("End")
    
'    Call Create_xml_data (Question, choices, answer)

End Function

Function Condition_data(str)

	If InStr (str, vbTab) > 0 Then
	
		str = Replace (str, vbTab, Chr(32))
		
	ElseIf InStr (str, "	") > 0 Then

		str = Replace (str, "	", Chr(32))
		
	ElseIf InStr (str, "  ") > 0 Then
	
		str = Replace (str, "  ", Chr (32)) 
		
	ElseIf InStr (str, "     ") > 0 Then

		str = Replace (str, "     ", Chr(32))
		
	ElseIf InStr (str, "     ") > 0 Then
	
		str = Replace (str, "     ", Chr(32))
		
	End If 
		
	Call create_conditioned_text_file(str)

End Function

Function create_conditioned_text_file(string_data)

    file_path = Split (pathToFile, ".", 2)
    
    file_name = file_path(0) & "_cond.txt"
    
    If fso.FileExists (file_name) = False Then

        Set xml_file = fso.CreateTextFile (file_name)
        
    Else 
    
        Set xml_file = fso.OpenTextFile (file_name, 8)
        
    End If
    
    If Data <> "End" Then 
    
        xml_file.WriteLine (string_data)
    
    End If    
    
    xml_file.Close
    
    Set xml_file = Nothing 

End Function