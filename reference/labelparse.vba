

'*********************************************************************************************
'***** Original Creation
'*****   Date: March 18, 2008
'*****   Author: Aubrey Sicotte
'*****   Purpose: Convert and parse surficial polygon data
'*****   Description: Loops through all labels in a chosen table, parses each label into
'*****                components and processes. Then each component and process is converted
'*****                (if necessary). Then Each label in concatenated back together and then
'*****                parsed back out into their individual types.
'***** Last Edit
'*****   Date:
'*****   Author:
'*****   Description:
'*****
'***** Required References (Tools->References):
'*****   Microsoft DAO 3.6 Object Library
'*****   Visual Basic For Applications
'*****   Microsoft Access 10.0 Object Library
'*****   OLE Automation
'*****   Microsoft ActiveX Data Objects 2.1 Library
'*********************************************************************************************

'Option Compare Database
Option Explicit

'** Global variables to temporarily store all the pieces of labels
Public label, comp_a, del_ab, comp_b, del_bc, comp_c, del_cd, comp_d, process As String
Public pos_ab, pos_bc, pos_cd As Integer
Public pre_partcov_a, partcov_a, bedrock_a, text1_a, text2_a, text3_a, mat_a, qual_a, st_a, express1_a, express2_a, express3_a, age_a As String
Public partcov_b, bedrock_b, text1_b, text2_b, text3_b, mat_b, qual_b, st_b, express1_b, express2_b, express3_b, age_b As String
Public partcov_c, bedrock_c, text1_c, text2_c, text3_c, mat_c, qual_c, st_c, express1_c, express2_c, express3_c, age_c As String
Public partcov_d, bedrock_d, text1_d, text2_d, text3_d, mat_d, qual_d, st_d, express1_d, express2_d, express3_d, age_d As String
Public process_a, proclass1a, proclass2a, process_b, proclass1b, proclass2b, process_c, proclass1c, proclass2c As String



Private Sub Form_Load()
'***** Sets all the Global Variables and sets up all the controls on the conversion form

    Dim db As DAO.Database
    Dim tb As DAO.TableDef
    '** Reset label (Global Variable)
    label = ""

    '** Reset all the Global Variables (except "label")
    Call RESET_GLOBALS

    '** open the current database
    Set db = CurrentDb

    '**Clear out and repopulate the combo-boxes on this form
    Me.COMP_CONV_TABLE.RowSource = ""
    Me.COMP_CONV_TABLE.Requery
    Me.PROC_CONV_TABLE.RowSource = ""
    Me.PROC_CONV_TABLE.Requery
    Me.TABLE_TO_CONV.RowSource = ""
    Me.TABLE_TO_CONV.Requery

    '** Loop through all the tables
    For Each tb In db.TableDefs
        '** Ignore all tables with "GDB" at the start (these are Geodatabase tables maintained by ESRI)
        If Left(tb.Name, 3) <> "GDB" Then
            '** List all tables ending in "CC" in the component conversion table combobox
            If Right(tb.Name, 2) = "CC" Then
                Me.COMP_CONV_TABLE.AddItem tb.Name
            '** List all tables ending in "CP" in the process conversion table combobox
            ElseIf Right(tb.Name, 2) = "CP" Then
                Me.PROC_CONV_TABLE.AddItem tb.Name
            '** List all other tables in the conversion table combobox
            Else
                Me.TABLE_TO_CONV.AddItem tb.Name
            End If
        End If
    Next tb

'**default values if you want them
'    Me.TABLE_TO_CONV.Value = "053_Hughes_Aishihik_polygons"
'    Me.COMP_CONV_TABLE.Value = "053_Hughes_Aishihik_Distinct_CC"
'    Me.PROC_CONV_TABLE.Value = "053_Hughes_Aishihik_Distinct_CP"

    Set db = Nothing
    Set tb = Nothing
End Sub

Private Sub CONVERTIT_Click()
'***** The main control routine. Runs the parsing and conversion process on each label (each record in the conversion table).

    Dim counterSP As Long
    Dim dbSP As DAO.Database
    Dim rstSP As DAO.Recordset
    Dim i As Integer
    Dim REPORT As String
    Dim tlabel, clabel As String

    '**Let the user know the conversion process has started
    DoCmd.Hourglass (True)

    '** Ensure that all three conversion tables are selected
    If Nz(Me.TABLE_TO_CONV.Value, "") = "" Or Nz(Me.PROC_CONV_TABLE.Value, "") = "" Or Nz(Me.COMP_CONV_TABLE.Value, "") = "" Then
        MsgBox "Please select tables for all three options"
        Exit Sub
    End If

'    Call CLEANER

    '** Reset all the Global Variables (except "label")
    Call RESET_GLOBALS

    '** open the current database
    Set dbSP = CurrentDb

    '** Open the conversion table
    Set rstSP = dbSP.OpenRecordset(Me.TABLE_TO_CONV.Value, dbOpenDynaset)

    '** If there are no records in the table, exit Sub
    If (rstSP.BOF = True And rstSP.EOF = True) Then
        Resume Exit_CONVERTIT
    End If

    '** Find out how many records there are in the source table
    If rstSP.RecordCount <> 0 Then
        rstSP.MoveLast
        counterSP = rstSP.RecordCount - 1
        rstSP.MoveFirst
    End If

    '** If there are no records in the source table, exit sub
    If Nz(Count, 0) = 0 Then
        Exit Sub
    End If

'    Call CLEANER

    '** loop through each record and perform all the conversion and parsing routines
        For i = 1 To counterSP
            label = Nz(rstSP!LABEL_NEW, "")

            '** ensure the label isn't empty
            If label <> "" And label <> "unknown" And label <> "unclassified" Then
'                If Left(label, 1) = Chr$(47) Then
'                    pre_partcov_a = Left(label, 1)
'                    label = Right(label, Len(label) - 1)
'                End If

                '** The next 7 lines are the core of this program
                Call PARSE_COMPONENTS
                Call PARSE_TYPES(True)
                Call CONVERT_COMPONENTS
                Call CONCAT
                Call PARSE_COMPONENTS
                Call PARSE_TYPES(False)
                Call PARSE_PSP

                '** set the recordset for editing
                rstSP.Edit

                '** concatenate all the components and processes for the LABEL_FNL field
                tlabel = CON_TLABEL()
                clabel = CON_LABEL()

                '** (for testing) test the converted label to see if it is as expected; if it isn't, set LABEL_TST = converted label
                If TEST_LABELS = False Then
                    rstSP![LABEL_TST] = clabel
                End If

                '** update all the appropriate fields from the values in the Global Variables
                rstSP![label_fnl] = tlabel
                rstSP![comp_a] = comp_a
                rstSP![RELATIONAB] = del_ab
                rstSP![comp_b] = comp_b
                rstSP![RELATIONBC] = del_bc
                rstSP![comp_c] = comp_c
                rstSP![RELATIONCD] = del_cd
                rstSP![comp_d] = comp_d
                rstSP![process] = process

                rstSP![partcov_a] = pre_partcov_a & partcov_a
                rstSP![bedrock_a] = bedrock_a
                rstSP![TEXTURE1_A] = text1_a
                rstSP![TEXTURE2_A] = text2_a
                rstSP![TEXTURE3_A] = text3_a
                rstSP![MATERIAL_A] = mat_a
                rstSP![QUALIFIERA] = qual_a
                rstSP![SUBTYPE_A] = st_a
                rstSP![EXPRSN1_A] = express1_a
                rstSP![EXPRSN2_A] = express2_a
                rstSP![EXPRSN3_A] = express3_a
                rstSP![age_a] = age_a

                rstSP![partcov_b] = partcov_b
                rstSP![bedrock_b] = bedrock_b
                rstSP![TEXTURE1_B] = text1_b
                rstSP![TEXTURE2_B] = text2_b
                rstSP![TEXTURE3_B] = text3_b
                rstSP![MATERIAL_B] = mat_b
                rstSP![QUALIFIERB] = qual_b
                rstSP![SUBTYPE_B] = st_b
                rstSP![EXPRSN1_B] = express1_b
                rstSP![EXPRSN2_B] = express2_b
                rstSP![EXPRSN3_B] = express3_b
                rstSP![age_b] = age_b

                rstSP![partcov_c] = partcov_c
                rstSP![bedrock_c] = bedrock_c
                rstSP![TEXTURE1_C] = text1_c
                rstSP![TEXTURE2_C] = text2_c
                rstSP![TEXTURE3_C] = text3_c
                rstSP![MATERIAL_C] = mat_c
                rstSP![QUALIFIERC] = qual_c
                rstSP![SUBTYPE_C] = st_c
                rstSP![EXPRSN1_C] = express1_c
                rstSP![EXPRSN2_C] = express2_c
                rstSP![EXPRSN3_C] = express3_c
                rstSP![age_c] = age_c

                rstSP![partcov_d] = partcov_d
                rstSP![bedrock_d] = bedrock_d
                rstSP![TEXTURE1_D] = text1_d
                rstSP![TEXTURE2_D] = text2_d
                rstSP![TEXTURE3_D] = text3_d
                rstSP![MATERIAL_D] = mat_d
                rstSP![QUALIFIERD] = qual_d
                rstSP![SUBTYPE_D] = st_d
                rstSP![EXPRSN1_D] = express1_d
                rstSP![EXPRSN2_D] = express2_d
                rstSP![EXPRSN3_D] = express3_d
                rstSP![age_d] = age_d

                rstSP![process_a] = process_a
                rstSP![proclass1a] = proclass1a
                rstSP![proclass2a] = proclass2a
                rstSP![process_b] = process_b
                rstSP![proclass1b] = proclass1b
                rstSP![proclass2b] = proclass2b
                rstSP![process_c] = process_c
                rstSP![proclass1c] = proclass1c
                rstSP![proclass2c] = proclass2c

                rstSP.Update

                pre_partcov_a = ""
            Else
                rstSP.Edit
                rstSP![label_fnl] = label
                rstSP.Update
            End If

            '** Reset label (Global Variable)
            label = ""

            '** Reset all the Global Variables (except "label")
            Call RESET_GLOBALS

            '** go to the next label (the next record in the conversion table)
            rstSP.MoveNext
        Next i

    '** close the conversion table
    rstSP.Close

    '** clear the recordset and database object (releases memory)
    Set rstSP = Nothing
    Set dbSP = Nothing

    '** clean the source table of all spaces and nulls
    'Call CLEAN_TABLE(Me.TABLE_TO_CONV.Value, counterSP)

    '** let the user know that the conversion is complete
    DoCmd.Hourglass False
    MsgBox "Conversion complete"

'** on a premature exit, clear the recordset and database objects
Exit_CONVERTIT:
    Set rstSP = Nothing
    Set dbSP = Nothing
    Exit Sub

'** on error, call the premature exit routine
Err_CONVERTIT:
    Resume Exit_CONVERTIT
End Sub


Private Sub PARSE_COMPONENTS()
'***** Parses the label into individual components
'*****   Note: Chr$(47) = "/"
'*****   Note: Chr$(46) = "."
'*****   Note: Chr$(92) = "\"

    Dim TC, TypeZ As String
    Dim i, countStart, lenlabel, posZ As Integer

    '** skip first char it is a "/"
    If Left(label, 1) = Chr$(47) Then
        countStart = 2
    Else
        countStart = 1
    End If

    '** find out the length of the label
    lenlabel = Len(label)

    '** loop to find the position of each delimeter
    For i = countStart To lenlabel
        '** reset the temporary position values
        posZ = 0
        TypeZ = ""

        '** assign each char (starting at the beginning of the string) to "TC"
        TC = Right(Left(label, i), 1)

        '** if TC is a delimeter, assign it to one of the position values
        If TC = Chr$(46) Or TC = Chr$(47) Or TC = Chr$(92) Then
            posZ = i
            Select Case TC
                Case Chr$(46)
                    TypeZ = Chr$(46)
                Case Chr$(92)
                    TypeZ = Chr$(92)
                Case Chr$(47)
                    TypeZ = Chr$(47)
                    If Right(Left(label, (i + 1)), 1) = Chr$(47) Then
                        TypeZ = TypeZ & Chr$(47)
                        i = (i + 1)
                    End If
            End Select
            If pos_ab = 0 Then
                pos_ab = posZ
                del_ab = TypeZ
            ElseIf pos_bc = 0 Then
                pos_bc = posZ
                del_bc = TypeZ
            ElseIf pos_cd = 0 Then
                pos_cd = posZ
                del_cd = TypeZ
            End If
        End If
    Next i

    '** if there are 3 delimeters, assign them accordingly
    If pos_cd > 0 Then
        If del_cd = Chr$(47) & Chr$(47) Then
            comp_d = Right(label, (lenlabel - (pos_cd + 1)))
        Else
            comp_d = Right(label, (lenlabel - pos_cd))
        End If
        If del_bc = Chr$(47) & Chr$(47) Then
            comp_c = Left(Right(label, lenlabel - (pos_bc + 1)), (pos_cd - pos_bc) - 2)
        Else
            comp_c = Left(Right(label, lenlabel - pos_bc), (pos_cd - pos_bc) - 1)
        End If

        If del_ab = Chr$(47) & Chr$(47) Then
            comp_b = Left(Right(label, lenlabel - (pos_ab + 1)), (pos_bc - pos_ab) - 2)
        Else
            comp_b = Left(Right(label, lenlabel - pos_ab), (pos_bc - pos_ab) - 1)
        End If
        comp_a = Left(label, (pos_ab - 1))

    '** if there are 2 delimeters, assign them accordingly
    ElseIf pos_bc > 0 Then
        If del_bc = Chr$(47) & Chr$(47) Then
            comp_c = Right(label, (lenlabel - (pos_bc + 1)))
        Else
            comp_c = Right(label, (lenlabel - pos_bc))
        End If

        If del_ab = Chr$(47) & Chr$(47) Then
            comp_b = Left(Right(label, (lenlabel - (pos_ab + 1))), (pos_bc - pos_ab) - 2)
        Else
            comp_b = Left(Right(label, (lenlabel - pos_ab)), (pos_bc - pos_ab) - 1)
        End If
        comp_a = Left(label, (pos_ab - 1))

    '** if there is 1 delimeter, assign the component and delimeter accordingly
    ElseIf pos_ab > 0 Then
        If del_ab = Chr$(47) & Chr$(47) Then
            comp_b = Right(label, (lenlabel - (pos_ab + 1)))
        Else
            comp_b = Right(label, (lenlabel - pos_ab))
        End If
        comp_a = Left(label, (pos_ab - 1))
    Else
        comp_a = label
    End If
End Sub

Private Sub CONVERT_COMPONENTS()
'***** loops through each record in the component conversion table to convert matching components
'*****   if the component = CC_Table.COMPS then component == CC_Table.COMPS_FINAL

    Dim counterCV As Integer
    Dim dbCV As DAO.Database
    Dim rstCV As DAO.Recordset
    Dim i As Integer
    Dim unit, label_fnl As String
    Dim done_conv As Boolean

    Set dbCV = CurrentDb

    '** Open the component conversion table
    Set rstCV = dbCV.OpenRecordset(Me.COMP_CONV_TABLE, dbOpenDynaset)

    '** if there are no records in the table, exit Sub
    If (rstCV.BOF = True And rstCV.EOF = True) Then
        Resume Exit_CONVERT_COMPONENTS
    End If

    '** find out how many records there are in the table
    If rstCV.RecordCount <> 0 Then
        rstCV.MoveLast
        counterCV = rstCV.RecordCount
        rstCV.MoveFirst
    End If

    '** loop through each record in the component conversion table; convert comp_a accordingly
    For i = 1 To counterCV

        '** set "finished conversion" flag to false (default)
        done_conv = False
        unit = rstCV!COMPS
        label_fnl = rstCV!COMPS_FINAL

        '** if the component = CC_Table.COMPS then component == CC_Table.COMPS_FINAL (P_PROCESSES strips off the processes)
        If comp_a = unit Then
            comp_a = P_PROCESSES(label_fnl, True)

            '** conversion complete
            done_conv = True
        End If

        '** if the conversion is complete, exit this for loop
        If done_conv = True Then
            Exit For
        End If

        '** move to the next record in table
        rstCV.MoveNext
    Next i

    '** move back to the start of the component conversion table
    rstCV.MoveFirst

    '** loop through each record in the component conversion table; convert comp_b accordingly
    For i = 1 To counterCV

        '** set "finished conversion" flag to false
        done_conv = False
        unit = rstCV!COMPS
        label_fnl = rstCV!COMPS_FINAL

        '** if the component = CC_Table.COMPS then component == CC_Table.COMPS_FINAL (P_PROCESSES strips off the processes)
        If comp_b = unit Then
            comp_b = P_PROCESSES(label_fnl, True)

            '** conversion complete
            done_conv = True
        End If

        '** if the conversion is complete, exit this for loop
        If done_conv = True Then
            Exit For
        End If

        '** move to the next record in table
        rstCV.MoveNext
    Next i

    '** move back to the start of the component conversion table
    rstCV.MoveFirst

    '** loop through each record in the component conversion table; convert comp_c accordingly
    For i = 1 To counterCV

        '** set "finished conversion" flag to false
        done_conv = False
        unit = rstCV!COMPS
        label_fnl = rstCV!COMPS_FINAL

        '** if the component = CC_Table.COMPS then component == CC_Table.COMPS_FINAL (P_PROCESSES strips off the processes)
        If comp_c = unit Then
            comp_c = P_PROCESSES(label_fnl, True)

            '** conversion complete
            done_conv = True
        End If
        If done_conv = True Then
            Exit For
        End If

        '** move to the next record in table
        rstCV.MoveNext
    Next i

    '** move back to the start of the component conversion table
    rstCV.MoveFirst

    '** loop through each record in the component conversion table; convert comp_d accordingly
    For i = 1 To counterCV

        '** set "finished conversion" flag to false
        done_conv = False
        unit = rstCV!COMPS
        label_fnl = rstCV!COMPS_FINAL

        '** if the component = CC_Table.COMPS then component == CC_Table.COMPS_FINAL (P_PROCESSES strips off the processes)
        If comp_d = unit Then
            comp_d = P_PROCESSES(label_fnl, True)

            '** conversion complete
            done_conv = True
        End If

        '** if the conversion is complete, exit this for loop
        If done_conv = True Then
            Exit For
        End If

        '** move to the next record in table
        rstCV.MoveNext
    Next i

    '** close recordset
    rstCV.Close

    '** release objects
    Set rstCV = Nothing
    Set dbCV = Nothing

'** on a premature exit, clear the recordset and database objects
Exit_CONVERT_COMPONENTS:
    Set rstCV = Nothing
    Set dbCV = Nothing
    Exit Sub

'** on error, call the premature exit routine
Err_CONVERT_COMPONENTS:
    Resume Exit_CONVERT_COMPONENTS
End Sub

Private Sub CONCAT()
'***** concatenates all the components and processes together

    label = comp_a & del_ab & comp_b & del_bc & comp_c & del_cd & comp_d
    If Len(process) > 0 Then
        label = label & "-" & process
    End If

    '** Reset all the Global Variables (except "label")
    Call RESET_GLOBALS
End Sub

Private Sub PARSE_TYPES(boolCONV As Boolean)
'***** directs the parsing of components and processes

    Dim tmp As String

    '** if comp_a exists, parse it
    If Len(comp_a) > 0 Then
        tmp = comp_a

            '** strip off the processes
            comp_a = P_PROCESSES(tmp, boolCONV)

            '** parse comp_a
            P_COMP ("a")
    End If

    '** if comp_b exists, parse it
    If Len(comp_b) > 0 Then
        tmp = comp_b

            '** strip off the processes
            comp_b = P_PROCESSES(tmp, boolCONV)

            '** parse comp_b
            P_COMP ("b")
    End If

    '** if comp_c exists, parse it
    If Len(comp_c) > 0 Then
        tmp = comp_c

            '** strip off the processes
            comp_c = P_PROCESSES(tmp, boolCONV)

            '** parse comp_c
            P_COMP ("c")
    End If

    '** if comp_d exists, parse it
    If Len(comp_d) > 0 Then
        tmp = comp_d

            '** strip off the processes
            comp_d = P_PROCESSES(tmp, boolCONV)

            '** parse comp_d
            P_COMP ("d")
    End If
End Sub

Private Sub P_COMP(comp As String)
'***** parses a component into its individual types
'*****   Note: each component is parsed into: partcov, bedrock, texture, material, qualifier, expression and age

    Dim component, eval, Ueval As String
    Dim fw, bw, mat_pos, i As Integer
    Dim partcov_t, bedrock_t, text1_t, text2_t, text3_t, mat_t, qual_t, express1_t, express2_t, express3_t, age_t As String
    Dim tmpL, tmpR As String

    '** select the appropriate component to work with
    If comp = "a" Then component = comp_a
    If comp = "b" Then component = comp_b
    If comp = "c" Then component = comp_c
    If comp = "d" Then component = comp_d

    '** fw = "forward" position marker for stepping forward through each char of the component
    fw = 1

    '** bw = "backward" position marker for stepping backward through each char of the component
    bw = Len(component)

    '** loop through each char and find the first upper case char = material
    For i = 0 To bw
        eval = (Left(Right(component, Len(component) - i), 1))
        If StrComp(eval, UCase(eval), vbBinaryCompare) = 0 Then

            '** assign material
            mat_t = eval

            '** assign material position
            mat_pos = i + 1
            Exit For
        End If
    Next i

    '** if the material is Rock ("R") Then the 2 chars before it are bedrock
    If mat_t = "R" And mat_pos = 3 Then
        bedrock_t = Left(component, 2)
    Else

        '** test to see if there are chars before material
        If mat_pos - fw > 0 Then

            '** parse pre-mat chars
            eval = Left(component, 1)
            If StrComp(eval, LCase(eval), vbBinaryCompare) <> 0 Then

                '** parse partcov
                partcov_t = Left(component, 1)
                fw = 2
            End If
        End If


        '** parse textures
        If mat_pos - fw > 0 Then
            text1_t = Right(Left(component, fw), 1)
            fw = fw + 1
        End If
        If mat_pos - fw > 0 Then
            text2_t = Right(Left(component, fw), 1)
            fw = fw + 1
        End If
        If mat_pos - fw > 0 Then
            text3_t = Right(Left(component, fw), 1)
        End If
    End If
    fw = mat_pos + 1
    bw = Len(component)

    '** parse qualifier
    If (bw - fw) >= 0 Then
        If StrComp(Right(Left(component, fw), 1), UCase(Right(Left(component, fw), 1)), vbBinaryCompare) = 0 And Right(Left(component, fw), 1) <> Chr$(60) And Right(Left(component, fw), 1) <> Chr$(62) Then
            If Right(Left(component, fw), 1) = "G" Or Right(Left(component, fw), 1) = "I" Or Right(Left(component, fw), 1) = "A" Then
                qual_t = Right(Left(component, fw), 1)
                fw = fw + 1
            End If
        End If
    End If

    '** parse ages
    If (bw - fw) >= 0 Then
        If StrComp(Left(Right(component, 1), 1), UCase(Left(Right(component, 1), 1)), vbBinaryCompare) = 0 Then
            If Left(Right(component, 2), 1) = Chr$(60) Or Left(Right(component, 2), 1) = Chr$(62) Then
                age_t = Right(component, 2)
            Else
                age_t = Right(component, 1)
            End If
        End If
    End If

    '** parse expressions
    If (bw - fw) >= 0 Then
        tmpL = Left(Right(component, fw), 1)
        tmpR = Right(Left(component, fw), 1)
        If StrComp(tmpR, LCase(tmpR), vbBinaryCompare) = 0 And tmpR <> Chr$(60) And tmpR <> Chr$(62) Then
            express1_t = Right(Left(component, fw), 1)
            fw = fw + 1
        End If
    End If
    If (bw - fw) >= 0 Then
        tmpL = Left(Right(component, fw), 1)
        tmpR = Right(Left(component, fw), 1)
        If StrComp(tmpR, LCase(tmpR), vbBinaryCompare) = 0 And tmpR <> Chr$(60) And tmpR <> Chr$(62) Then
            express2_t = Right(Left(component, fw), 1)
            fw = fw + 1
        End If
    End If
    If (bw - fw) >= 0 Then
        tmpL = Left(Right(component, fw), 1)
        tmpR = Right(Left(component, fw), 1)
        If StrComp(tmpR, LCase(tmpR), vbBinaryCompare) = 0 And tmpR <> Chr$(60) And tmpR <> Chr$(62) Then
            express3_t = Right(Left(component, fw), 1)
        End If
    End If

    '** assign parsed values to Global Variables
    If comp = "a" Then
        partcov_a = partcov_t
        bedrock_a = bedrock_t
        text1_a = text1_t
        text2_a = text2_t
        text3_a = text3_t
        mat_a = mat_t
        qual_a = qual_t
        express1_a = express1_t
        express2_a = express2_t
        express3_a = express3_t
        age_a = age_t
    ElseIf comp = "b" Then
        partcov_b = partcov_t
        bedrock_b = bedrock_t
        text1_b = text1_t
        text2_b = text2_t
        text3_b = text3_t
        mat_b = mat_t
        qual_b = qual_t
        express1_b = express1_t
        express2_b = express2_t
        express3_b = express3_t
        age_b = age_t
    ElseIf comp = "c" Then
        partcov_c = partcov_t
        bedrock_c = bedrock_t
        text1_c = text1_t
        text2_c = text2_t
        text3_c = text3_t
        mat_c = mat_t
        qual_c = qual_t
        express1_c = express1_t
        express2_c = express2_t
        express3_c = express3_t
        age_c = age_t
    ElseIf comp = "d" Then
        partcov_d = partcov_t
        bedrock_d = bedrock_t
        text1_d = text1_t
        text2_d = text2_t
        text3_d = text3_t
        mat_d = mat_t
        qual_d = qual_t
        express1_d = express1_t
        express2_d = express2_t
        express3_d = express3_t
        age_d = age_t
    End If
End Sub

Private Function P_PROCESSES(component As String, boolCONV As Boolean) As String
'***** directs the parsing of processes

    Dim hasProc, i, temp_numP, numP, iPSP As Integer
    Dim PROC, comp, eP, eSP1, eSP2, eSP3 As String
    Dim t, tmp As Boolean

    '**  set "component has a process" flag to false (default)
    comp = component

    '** find the starting point of processes
    hasProc = InStr(comp, "-")

    '** if there are processes, parse them
    If hasProc > 0 Then

        '** separate the processes from the component
        PROC = Right(comp, Len(comp) - hasProc)
        comp = Left(comp, hasProc - 1)
        '** find out how many processes there are
        For i = 1 To Len(PROC)
            If StrComp(Right(Left(PROC, i), 1), UCase(Right(Left(PROC, i), 1)), vbBinaryCompare) = 0 Then
                numP = numP + 1
            End If
        Next i

        '** parse out and then insert each process into global process variable (max = AaaaBbbbCccc)
        For temp_numP = 1 To numP

            '** examine each character at a time, extract 'Aa' combinations, insert them into global variable
            eP = Left(PROC, 1)
            iPSP = 1
            If Len(PROC) > 1 Then
                If StrComp(Right(Left(PROC, 2), 1), LCase(Right(Left(PROC, 2), 1)), vbBinaryCompare) = 0 Then
                    eSP1 = Right(Left(PROC, 2), 1)
                    iPSP = 2
                    If Len(PROC) > 2 Then
                        If StrComp(Right(Left(PROC, 3), 1), LCase(Right(Left(PROC, 3), 1)), vbBinaryCompare) = 0 Then
                            eSP2 = Right(Left(PROC, 3), 1)
                            iPSP = 3
                            If Len(PROC) > 3 Then
                                If StrComp(Right(Left(PROC, 4), 1), LCase(Right(Left(PROC, 4), 1)), vbBinaryCompare) = 0 Then
                                    eSP3 = Right(Left(PROC, 4), 1)
                                    iPSP = 4
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            '** if "process needs to be converted" is false, then just insert the process
            If boolCONV = False Then
                If iPSP = 1 Then
                    INSERT_PROCESS (Nz(eP, ""))
                Else
                    If iPSP > 1 Then
                        INSERT_PROCESS (Nz((eP & eSP1), ""))
                    End If
                    If iPSP > 2 Then
                        INSERT_PROCESS (Nz((eP & eSP2), ""))
                    End If
                    If iPSP > 3 Then
                        INSERT_PROCESS (Nz((eP & eSP3), ""))
                    End If
                End If

            '** if "process needs to be converted" is true, then convert the process and then insert it
            Else
                If iPSP = 1 Then
                    INSERT_PROCESS (CONVERT_PSP(Nz(eP, "")))
                Else
                    If iPSP > 1 Then
                        INSERT_PROCESS (CONVERT_PSP(Nz((eP & eSP1), "")))
                    End If
                    If iPSP > 2 Then
                        INSERT_PROCESS (CONVERT_PSP(Nz((eP & eSP2), "")))
                    End If
                    If iPSP > 3 Then
                        INSERT_PROCESS (CONVERT_PSP(Nz((eP & eSP3), "")))
                    End If
                End If
            End If
            If Len(PROC) > 0 Then PROC = Right(PROC, Len(PROC) - iPSP)
        Next temp_numP
    End If

    '** return the process without any processes
    P_PROCESSES = comp
End Function

Private Function CONVERT_PSP(PSP As String) As String
'***** loops through each record in the process conversion table to convert matching processes
'*****   if the component = CP_Table.PROCESSES then component == CC_Table.PROCESSES_FINAL

    Dim counterCV As Integer
    Dim dbCV As DAO.Database
    Dim rstCV As DAO.Recordset
    Dim i As Integer
    Dim PROC_START, PROC_END, strCONV_PSP As String
    Dim done_conv As Boolean

    Set dbCV = CurrentDb

    '** Open the process conversion table
    Set rstCV = dbCV.OpenRecordset(Me.PROC_CONV_TABLE, dbOpenDynaset)

    '** if there are no records in the table, exit this function
    If (rstCV.BOF = True And rstCV.EOF = True) Then
        Resume Exit_CONVERT_PSP
    End If

    '** find out how many records there are in the table
    If rstCV.RecordCount <> 0 Then
        rstCV.MoveLast
        counterCV = rstCV.RecordCount

        '** go to the first record
        rstCV.MoveFirst
    End If

    '** loop through each record in the process conversion table
    For i = 1 To counterCV

        '** set "finished conversion" flag to false (default)
        done_conv = False

        '** get CP_Table.PROCESS
        PROC_START = rstCV!PROCESSES

        '** get CP_Table.PROCESS_FINAL
        PROC_END = rstCV!PROCESSES_FINAL

        '** if the process = CC_Table.PROCESSES then process == CC_Table.PROCESSES_FINAL
        If StrComp(PSP, PROC_START, vbTextCompare) = 0 Then
            strCONV_PSP = PROC_END
            Exit For
        Else
            strCONV_PSP = PSP
        End If

        '** move to next record in table
        rstCV.MoveNext
    Next i

    '** return converted process
    CONVERT_PSP = strCONV_PSP
    rstCV.Close

    '** release objects
    Set rstCV = Nothing
    Set dbCV = Nothing

'** on a premature exit, clear the recordset and database objects
Exit_CONVERT_PSP:
    Set rstCV = Nothing
    Set dbCV = Nothing
    Exit Function

'** on error, call the premature exit routine
Err_CONVERT_PSP:
    Resume Exit_CONVERT_PSP
    Exit Function
End Function

Private Sub INSERT_PROCESS(i_proc As String)
'***** inserts a new process into the process Global Variable

    Dim newP, newSP, oldProc, p, pA, pB As String
    Dim boolNoAdd, procExists As Boolean
    Dim lenNewProc, oldP_pos, i, numSP As Integer
    Dim tempL, tempR, tempM As String

    '** find the length of the process
    lenNewProc = Len(i_proc)

    '** set "insert this process into process Global Variable" flag to false (default)
    boolNoAdd = False

    '** set "process Global Variable exists" flag to false (default)
    procExists = False

    '** set number of sub-processes to 0 (default)
    numSP = 0

    '** for storing sub-processes
    pA = ""
    pB = ""

    Select Case lenNewProc

        '** if there is no process, do nothing
        Case 0
            process = i_proc

        '** if the new process is only one char & if it isn't in the process Global Variable, add it to the process Global Variable
        Case 1
            If InStr(process, i_proc) = 0 Then
               process = process & i_proc
            End If

        '** if the new process is two chars & if they aren't in the process Global Variable, add them to the process Global Variable
        Case 2

            '** set the process variable
            newP = Left(i_proc, 1)

            '** set the sub-process variable
            newSP = Right(i_proc, 1)

            '** process isn't in process Global Variable so add it
            If InStr(process, newP) = 0 Then
                process = process & i_proc

            '** process is in process Global Variable, check sub-process
            Else
                oldProc = process

                '** find position of process in process Global Variable
                oldP_pos = InStr(process, newP)

                '** look for process in process Global Variable
                For i = 1 To Len(oldProc)

                    '** if the process does exist in process Global Variable then look for sub-processes
                    If Right(Left(oldProc, i), 1) = Left(i_proc, 1) Then
                        procExists = True
                        p = Right(Left(oldProc, i), 1)
                        If (i) < Len(oldProc) Then

                            '** look for sub-process to add
                            If StrComp(Right(Left(oldProc, i + 1), 1), LCase(Right(Left(oldProc, i + 1), 1)), vbBinaryCompare) = 0 Then

                                '** 1 sub-process to add
                                numSP = 1

                                '** 1st sub-process to add
                                pA = Right(Left(oldProc, i + 1), 1)
                            End If
                        End If
                        If numSP = 1 And (i + 1) < Len(oldProc) Then

                            '** look for sub-process to add
                            If StrComp(Right(Left(oldProc, i + 2), 1), LCase(Right(Left(oldProc, i + 2), 1)), vbBinaryCompare) = 0 Then

                                '** 2 sub-processes to add
                                numSP = 2

                                '** 2nd sub-process to add
                                pB = Right(Left(oldProc, i + 2), 1)
                            End If
                        End If
                        Exit For
                    End If
                Next i

                '** insert the process(es) into process Global Variable
                Select Case numSP

                    '** if there are no sub-processes, insert the process (if applicable)
                    Case 0
                        tempL = Left(oldProc, oldP_pos)
                        tempM = Right(i_proc, 1)
                        If Len(oldProc) > oldP_pos Then
                            tempR = Right(oldProc, Len(oldProc) - (oldP_pos + 1))
                        Else
                            tempR = ""
                        End If
                        process = tempL & tempM & tempR

                    '** if there is 1 sub-process, insert the process and sub-process (if applicable)
                    Case 1
                        If StrComp(pA, newSP, vbBinaryCompare) <> 0 Then
                            tempL = Left(oldProc, oldP_pos + 1)
                            tempM = Right(i_proc, 1)
                            tempR = Right(oldProc, (Len(oldProc) - (oldP_pos + 1)))
                            process = tempL & tempM & tempR
                        End If

                    '** if there are 2 sub-processes, insert the process and sub-process(es) (if applicable)
                    Case 2
                        If StrComp(pA, newSP, vbBinaryCompare) <> 0 And StrComp(pB, newSP, vbBinaryCompare) <> 0 Then
                            tempL = Left(oldProc, oldP_pos + 2)
                            tempM = Right(i_proc, 1)
                            tempR = Right(oldProc, (Len(oldProc) - (oldP_pos + 2)))
                            process = tempL & tempM & tempR
                        End If
                End Select
            End If
    End Select
End Sub

Private Sub PARSE_PSP()
'***** parses processes and sub-processes into their individual types
'*****   Note: each process can have up to 3 processes with 2 sub-processes each
'*****   Note: each process is then parsed into process, proclass1, proclass2

    Dim i, j, numProc As Integer
    Dim tprocess As String

    '** counters for stepping through process by char
    j = 1
    i = 1
    tprocess = process

    '** make sure that the process isn't an empty string
    If Len(tprocess) = 0 Then
        Exit Sub
    End If

    '** count the number of processes
    For j = 1 To Len(tprocess)
        If StrComp(Right(Left(tprocess, j), 1), UCase(Right(Left(tprocess, j), 1)), vbBinaryCompare) = 0 Then
            numProc = numProc + 1
        End If
    Next j

    '** parse the 1st process/sub-processes set into types
    If numProc > 0 Then
        process_a = Left(tprocess, 1)
        If Len(tprocess) > i Then
            If StrComp(Right(Left(tprocess, (i + 1)), 1), LCase(Right(Left(tprocess, (i + 1)), 1)), vbBinaryCompare) = 0 Then
                i = i + 1
                proclass1a = Right(Left(tprocess, i), 1)
                If Len(tprocess) > i Then
                    If Len(proclass1a) > 0 And StrComp(Right(Left(tprocess, i + 1), 1), LCase(Right(Left(tprocess, i + 1), 1)), vbBinaryCompare) = 0 Then
                        i = i + 1
                        proclass2a = Right(Left(tprocess, i), 1)
                    End If
                End If
            End If
        End If

        '** trim off the process/sub-processes set that was just parsed
        tprocess = Right(process, Len(tprocess) - i)
    End If

    '** parse the 2nd process/sub-processes set into types
    i = 1
    If numProc > 1 Then
        process_b = Left(tprocess, 1)
        If Len(tprocess) > i Then
            If StrComp(Right(Left(tprocess, (i + 1)), 1), LCase(Right(Left(tprocess, (i + 1)), 1)), vbBinaryCompare) = 0 Then
                i = i + 1
                proclass1b = Right(Left(tprocess, i), 1)
                If Len(tprocess) > i Then
                    If Len(proclass1b) > 0 And StrComp(Right(Left(tprocess, i + 1), 1), LCase(Right(Left(tprocess, i + 1), 1)), vbBinaryCompare) = 0 Then
                        i = i + 1
                        proclass2b = Right(Left(tprocess, i), 1)
                    End If
                End If
            End If
        End If

        '** trim off the process/sub-processes set that was just parsed
        tprocess = Right(process, Len(tprocess) - i)
    End If

    '** parse the 3rd process/sub-processes set into types
    i = 1
    If numProc > 2 Then
        process_c = Left(tprocess, 1)
        If Len(tprocess) > i Then
            If StrComp(Right(Left(tprocess, (i + 1)), 1), LCase(Right(Left(tprocess, (i + 1)), 1)), vbBinaryCompare) = 0 Then
                i = i + 1
                proclass1c = Right(Left(tprocess, i), 1)
                If Len(tprocess) > i Then
                    If Len(proclass1c) > 0 And StrComp(Right(Left(tprocess, i + 1), 1), LCase(Right(Left(tprocess, i + 1), 1)), vbBinaryCompare) = 0 Then
                        i = i + 1
                        proclass2c = Right(Left(tprocess, i), 1)
                    End If
                End If
            End If
        End If

        '** trim off the process/sub-processes set that was just parsed
        tprocess = Right(process, Len(tprocess) - i)
    End If

End Sub

Function fExistTable(strTableName As String) As Boolean
'***** function to find out if a table exists

Dim db As DAO.Database
Dim i As Integer
    Set db = DBEngine.Workspaces(0).Databases(0)
    fExistTable = False
    db.TableDefs.Refresh
    For i = 0 To db.TableDefs.Count - 1
        If strTableName = db.TableDefs(i).Name Then
            '** Table Exists
            fExistTable = True
            Exit For
        End If
    Next i
    Set db = Nothing
End Function

Private Function TEST_LABELS() As Boolean
'***** compare (concat components) and (concat individual types) to check if they are the same
'*****  if they are the same: return "True"
'*****  if they are not the same: return "False"

    Dim flag As Boolean
    Dim clabel, tlabel As String
    Dim i As Integer

    '** concatenate each individual type together
    clabel = CON_LABEL

    '** concatenate each component together
    tlabel = CON_TLABEL

    '** set "they are the same" flag to True
    flag = True

    '** check to see if they are the same
    If Len(clabel) = Len(tlabel) Then
        For i = 1 To Len(clabel)
            If StrComp(Right(Left(clabel, i), 1), Right(Left(tlabel, i), 1), vbBinaryCompare) <> 0 Then
                flag = False
            End If
        Next
    Else
        flag = False
    End If

    '** return the results (True or False)
    TEST_LABELS = flag
End Function

Private Function CON_LABEL() As String
'***** concatenate each individual type together

    Dim tmp, tmpA, tmpB, tmpC, tmpD, tmpP As String

    tmpA = partcov_a & bedrock_a & text1_a & text2_a & text3_a & mat_a & qual_a & st_a & express1_a & express2_a & express3_a & age_a
    tmpB = partcov_b & bedrock_b & text1_b & text2_b & text3_b & mat_b & qual_b & st_b & express1_b & express2_b & express3_b & age_b
    tmpC = partcov_c & bedrock_c & text1_c & text2_c & text3_c & mat_c & qual_c & st_c & express1_c & express2_c & express3_c & age_c
    tmpD = partcov_d & bedrock_d & text1_d & text2_d & text3_d & mat_d & qual_d & st_d & express1_d & express2_d & express3_d & age_d
    tmpP = process_a & proclass1a & proclass2a & process_b & proclass1b & proclass2b & process_c & proclass1c & proclass2c
    tmp = tmpA & del_ab & tmpB & del_bc & tmpC & del_cd & tmpD
    If Len(tmpP) > 0 Then
        tmp = tmp & "-" & tmpP
    End If
    CON_LABEL = tmp
End Function

Private Function CON_TLABEL() As String
'***** concatenate each component together

    Dim tmp As String

    tmp = comp_a & del_ab & comp_b & del_bc & comp_c & del_cd & comp_d
    If Len(process) > 0 Then
        tmp = tmp & "-" & process
    End If
    CON_TLABEL = tmp
End Function

'Private Sub CLEAN_TABLE(tablename As String, counterSP As Long)
''***** Clean the table of all spaces and nulls
'
'    Dim i, j, numFields As Integer
'    Dim dbSP As DAO.Database
'    Dim rstSP As DAO.Recordset
'    Dim tbIndex As Integer
'    Dim tds As DAO.TableDefs
'    Dim td As DAO.TableDef
'
'    '** open the current database
'    Set dbSP = CurrentDb
'    Set tds = dbSP.TableDefs
'
'    '** Open the conversion table
'    Set rstSP = dbSP.OpenRecordset(tablename, dbOpenDynaset)
'
'    '** find out how many fields there are
'    numFields = rstSP.Fields.Count
'
'    '** go to the first record
'    rstSP.MoveFirst
'
'    '** change all nulls and spaces to empty strings
'    For i = 1 To counterSP
'        For j = 1 To numFields - 1
'            If rstSP.Fields(j).Type = dbText Then
'                If rstSP.Fields(j) = Null Or rstSP.Fields(j) = " " Or rstSP.Fields(j) = "  " Then
'                    '** set the recordset for editing
'                    rstSP.Edit
'                    rstSP.Fields(j) = ""
'                    rstSP.Update
'                End If
'            End If
'        Next j
'        rstSP.MoveNext
'    Next i
'
'
''** on a premature exit, clear the recordset and database objects
'Exit_CONVERTIT:
'    Set rstSP = Nothing
'    Set dbSP = Nothing
'    Exit Sub
'
''** on error, call the premature exit routine
'Err_CONVERTIT:
'    Resume Exit_CONVERTIT
'End Sub

Private Sub RESET_GLOBALS()
'***** resets all the Global Variables (except label)

    comp_a = ""
    del_ab = ""
    comp_b = ""
    del_bc = ""
    comp_c = ""
    del_cd = ""
    comp_d = ""
    process = ""

    pos_ab = 0
    pos_bc = 0
    pos_cd = 0

    partcov_a = ""
    bedrock_a = ""
    text1_a = ""
    text2_a = ""
    text3_a = ""
    mat_a = ""
    qual_a = ""
    st_a = ""
    express1_a = ""
    express2_a = ""
    express3_a = ""
    age_a = ""

    partcov_b = ""
    bedrock_b = ""
    text1_b = ""
    text2_b = ""
    text3_b = ""
    mat_b = ""
    qual_b = ""
    st_b = ""
    express1_b = ""
    express2_b = ""
    express3_b = ""
    age_b = ""

    partcov_c = ""
    bedrock_c = ""
    text1_c = ""
    text2_c = ""
    text3_c = ""
    mat_c = ""
    qual_c = ""
    st_c = ""
    express1_c = ""
    express2_c = ""
    express3_c = ""
    age_c = ""

    partcov_d = ""
    bedrock_d = ""
    text1_d = ""
    text2_d = ""
    text3_d = ""
    mat_d = ""
    qual_d = ""
    st_d = ""
    express1_d = ""
    express2_d = ""
    express3_d = ""
    age_d = ""

    process_a = ""
    proclass1a = ""
    proclass2a = ""
    process_b = ""
    proclass1b = ""
    proclass2b = ""
    process_c = ""
    proclass1c = ""
    proclass2c = ""
End Sub



