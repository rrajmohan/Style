Dim fso, objShell, objSource, item, zip_contnts

Dim i, j, k, l

i = 0

j = 0

k = 0

l = 1

ReDim Question(0)
    
ReDim choices(0)
    
ReDim answer(0)

Set fso = WScript.CreateObject("Scripting.Filesystemobject")

Set objShell = CreateObject("Shell.Application")

pathToFile = "C:\Users\Sujatha\Desktop\Raj\cs9.txt"

Chapter_name = "PRESENTATION"

Chapter_number = "09"

If fso.FileExists (pathToFile) = True Then

    Call Read_Data_From_Text_File(pathToFile)
    
Else 

    MsgBox "File does not exist Dummy !!!"
    
End If 

Function Read_Data_From_Text_File(filepath)

    Set open_file = fso.OpenTextFile (filepath, 1)

    Do Until open_file.AtEndOfStream
    
        Data_text = open_file.ReadLine
        
        Call Parse_data(Data_text)
    
    Loop
        
    open_file.Close
    
    Call Create_xml_data (Question, choices, answer)

End Function

Function Parse_data(Data_stream)

    If InStr (Data_stream, l & ".") > 0 Then
    
        ReDim Preserve Question(i)
        
        question_text = Split (Data_stream, l & ".", 2)
        
        Question(i) = question_text (1)
        
        i = i + 1
        
        l = l + 1
        
    ElseIf InStr (Data_stream, "a)") > o Then 
    
        ReDim Preserve choices(j)
        
        choices(j) = Data_stream
        
        j = j + 1
        
    ElseIf InStr (Data_stream, "answer") > 0 Then
    
        right_answer = Split (Data_stream, " ", 3)
        
        ReDim Preserve answer(k)
            
        answer(k) = Replace (right_answer (2), " ", "")
        
        k = k + 1
        
    End If

End Function

Function Create_xml_data (Question, choices, answer)

    Dim id, Question_number
    
    Volume_number = "01"
    
    xml_header = "<?xml version=" & Chr (34) & "1.0" & Chr (34) & " encoding=" & Chr (34) & "utf-8" & Chr (34) & "?> " & vbCrLf & "<chapter-questions>"
    
    Call Create_xml_file (xml_header)
    
    Chapter_name_tag = "<chapter name = " & Chr (34) & Chapter_name & Chr (34) & " number=" & Chr (34) & Chapter_number & chr (34) & ">"
    
    Call Create_xml_file (Chapter_name_tag)
    
    Question_number = 1

    For x = LBound(Question) To UBound(Question)
    
        Call Create_xml_file ("<question>")
        
        length = Len (Question_number)
        
        Question_number_tag = String ((4-length), "0") & Question_number
        
        id = "TNHSCS" & Volume_number & Chapter_number & Question_number_tag
    
        qid = "<questionId>" & id & "</questionId>"
        
        Call Create_xml_file(qid)
        
        Question_tag = "<question-text>" & Trim (Question(x)) & "</question-text>"
        
        Call Create_xml_file(Question_tag)
            
        answer_choice  = Split (choices(x), "b)", 2)
        
        answer_choice_1 = Replace (answer_choice(0), "a)", "")
        
        answer_choice_1 = Trim (answer_choice_1)

        answer_choice2 = Split (answer_choice(1), "c)", 2)
        
        answer_choice_2 = answer_choice2(0)
        
        answer_choice_2 = Trim (answer_choice_2)
        
        answer_choice3 = Split (answer_choice2(1), "d)", 2)
        
        answer_choice_3 = answer_choice3(0)
        
        answer_choice_3 = Trim (answer_choice_3)
        
        answer_choice_4 = answer_choice3(1)
        
        answer_choice_4 = Trim (answer_choice_4)
        
        Choices_tag = "<answer0>" & answer_choice_1 & "</answer0>" & vbCrLf & "<answer1>" & answer_choice_2 & "</answer1>" &  vbCrLf & "<answer2>" & answer_choice_3 & "</answer2>" &  vbCrLf & "<answer3>" & answer_choice_4 & "</answer3>"

        Call Create_xml_file(Choices_tag)
                
        grouping_tag = "<grouping>" & Chapter_number & "</grouping>"
        
        Call Create_xml_file(grouping_tag)
        
        weightage_tag = "<weightage>8</weightage>"
        
        Call Create_xml_file(weightage_tag)
        
        Correct_answer = answer(x)
        
        If Correct_answer = "a" Then
        
            Correct_answer_tag = "<correctAnswer>" & answer_choice_1 & "</correctAnswer>"
            
            Call Create_xml_file(Correct_answer_tag)
            
        ElseIf Correct_answer = "b" Then
        
            Correct_answer_tag = "<correctAnswer>" & answer_choice_2 & "</correctAnswer>"
            
            Call Create_xml_file(Correct_answer_tag)
            
        ElseIf Correct_answer = "c" Then
        
            Correct_answer_tag = "<correctAnswer>" & answer_choice_3 & "</correctAnswer>"
            
            Call Create_xml_file(Correct_answer_tag)
            
        ElseIf Correct_answer = "d" Then
        
            Correct_answer_tag = "<correctAnswer>" & answer_choice_4 & "</correctAnswer>"
            
            Call Create_xml_file(Correct_answer_tag)
            
        End If

        Call Create_xml_file("</question>")
        
        Question_number = Question_number + 1
        
    Next
    
    Call Create_xml_file("</chapter>")
    
    Call Create_xml_file("</chapter-questions>")
    
    Call Create_xml_file ("End")

End Function

Function Create_xml_file(data)

    xml_file_path = Split (pathToFile, ".", 2)
    
    xml_file_name = xml_file_path(0) & ".xml"

'	xml_file_name = "C:\Users\Sujatha\Desktop\Raj\cs.xml"
    
    If fso.FileExists (xml_file_name) = False Then

        Set xml_file = fso.CreateTextFile (xml_file_name)
        
    Else 
    
        Set xml_file = fso.OpenTextFile (xml_file_name, 8)
        
    End If
    
    If Data <> "End" Then 
    
        xml_file.WriteLine (Data)
        
    Elseif Data = "End" Then 
        
        xml_file.Close
        
    End If
    
    Set xml_file = Nothing 

End Function