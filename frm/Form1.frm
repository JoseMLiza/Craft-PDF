VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "cCraft-PDF"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   129
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   216
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMain 
      Caption         =   "Header - Fotter"
      Height          =   615
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton CmdMain 
      Caption         =   "Demo 1"
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Sub Form_Load()
   
    '-- GLOBAL: Enable Zlib to compress the content (Optional)
    LoadLibrary App.Path & "\bin\zlib.dll"
    
End Sub

Private Sub CmdMain_Click(Index As Integer)
    Select Case Index
        Case 0: Call Demo1
        Case 1: Call Demo2
    End Select
End Sub



Private Sub Demo1()

    Dim Obj As New cCraft
    With Obj
        
        '   DEFAULT
        '-----------------------------
        '   Document : A4 Page Size
        '   Font     : Helvetica
        '   FontSize : 10
        '-----------------------------
        
        .DrawRectangle 40, 30, .PageWidth - 80, 40, &H544336
        .SetFont HELVETICA_BOLD, 16, vbWhite
        .DrawText "PDF - VisualBasic 6", 0, 40, T_CENTER
        
        .DrawImage "img1", 45, 80, 80, 80
        .SetFont HELVETICA, 10, &H222222
        .DrawText FILE("info.txt"), 135, 85, T_JUSTIFIED, .PageWidth - 176, 3
        .DrawLine 40, 180, .PageWidth - 40, 180, vbRed, 0.4
        
        .SetFont HELVETICA_BOLD
        .DrawText "TIPOGRAFIA", 40, 190
        .SetFont HELVETICA
        .DrawText "Soporte para las 14 fuentes estándar del formato PDF.", 40, 208
        .DrawLine 40, 230, .PageWidth - 40, 230, vbRed, 0.5, "8 2" ' 8pt line - 3pt space
        
        .SetFont HELVETICA, , &HBC8F49
        .DrawText "Helvetica", 50, 240
        .SetFont HELVETICA_BOLD
        .DrawText "Helvetica-Bold", 50, 260
        .SetFont HELVETICA_OBLIQUE
        .DrawText "Helvetica-Oblique", 50, 280
        .SetFont HELVETICA_BOLD_OBLIQUE
        .DrawText "Helvetica-BoldOblique", 50, 300
        
        .SetFont TIMES_ROMAN, , &H32A944
        .DrawText "Times-Roman", 175, -240
        .SetFont TIMES_BOLD
        .DrawText "Times-Bold", 175, -260
        .SetFont TIMES_ITALIC
        .DrawText "Times-Italic", 175, -280
        .SetFont TIMES_BOLD_ITALIC
        .DrawText "Times-BoldItalic", 175, -300
        
        .SetFont COURIER, , &H1E3CE7
        .DrawText "Courier", 300, 240
        .SetFont COURIER_BOLD
        .DrawText "Courier-Bold", 300, 260
        .SetFont COURIER_OBLIQUE
        .DrawText "Courier-Oblique", 300, 280
        .SetFont COURIER_BOLD_OBLIQUE
        .DrawText "Courier-BoldOblique", 300, 300
        .DrawLine 40, 320, .PageWidth - 40, 320, vbBlue, 0.4
        
        .SetFont SYMBOL, , &HC21A7D
        .DrawText "abcYWD", 480, 240
        .SetFont ZAPF_DINGBATS
        .DrawText "abcdABC", 480, 260
        
        .SetFont HELVETICA, , &H222222
        .DrawText "Dibujo de lineas y rectángulos, con diferentes estilos de linea.", 40, 335
        
        .DrawRectangle 40, 360, 80, 50, , vbBlack, 1, Radius:=5             ' Only Border
        .DrawRectangle 140, 360, 80, 50, &H66B6FF, Radius:=5                ' Only Fill
        .DrawRectangle 240, 360, 80, 50, &H66B6FF, vbBlack, 1, Radius:=5    ' Border and Fill
        .DrawRectangle 340, 360, 80, 50, , vbBlack, 1, "3", 5               ' Only Border Dash (3pt line - 3pt space)
        .DrawRectangle 440, 360, 80, 50, &H8EF053, vbBlack, 1, "5 3", 5     ' Border Dash and Fill (5pt line - 3pt space)
        
        .DrawText "Soporte para imágenes con canal alpha.", 40, 450
        .DrawImage "img2", 40, 475, 50, 50
        .DrawImage "img3", 120, 475, 50, 50
        .DrawImage "img4", 200, 475, 50, 50
        .DrawImage "img5", 280, 475, 50, 50
        
        
        '-- Images --
        .AddImage LoadResData(101, "PNG"), "img1"
        .AddImage LoadResPicture(102, vbResBitmap), "img2"
        .AddImage Me.Icon, "img3"
        .AddImage App.Path & "\img\32bpp_2x2.png", "img4"
        .AddImage App.Path & "\img\001.bmp", "img5"
        
        
        '-- New Page (Table) --
        .AddPage
        .SetFont HELVETICA, 12
        .DrawText "Página 02", 0, 40, T_CENTER
        .DrawLine 40, 60, .PageWidth - 40, 60, vbRed, 0.1
        .DrawRectangle 40, 80, .PageWidth - 80, 20, &H808080
        
        Dim ColW As Variant:  ColW = Array(60, 150, 150, 130)
        Dim Cols  As Variant: Cols = Array("ID", "Nombre", "Apellido", "Ocupación")
        Dim x As Single, y As Single
        Dim i As Long, c As Long
    
        x = 40
        y = 82
        
        .SetFont HELVETICA_BOLD, 10, vbWhite
        For i = 0 To UBound(Cols)
            .DrawText Cols(i), x + 2, y + 2, IIf(i = 0, T_CENTER, 0), ColW(i)
            x = x + ColW(i)
        Next
        y = y + 20
        
        For i = 0 To 25
        
            If Not (i Mod 2) = 0 Then .DrawRectangle 40, y, .PageWidth - 80, 20, &HEFEFEF
            
            x = 40
            For c = 0 To UBound(Cols)
                Select Case c
                    Case 0
                        .SetFont COURIER_BOLD, 9, &H4343FF
                        .DrawText Format(i, "0000"), x + 2, (y + 4), T_CENTER, ColW(0)
                    Case 1
                        .SetFont HELVETICA, , &H222222
                        .DrawImage "img2", x + 1, -y - 2, 15, 15
                        .DrawText RandomName, x + 20, (y + 4)
                    Case 2
                        .DrawText RandomLastName & " " & RandomLastName, x + 2, (y + 4)
                    Case 3
                        .DrawText RandomJob, x + 2, (y + 4)
                End Select
                x = x + ColW(c)
            Next
            y = y + 20
        Next
        
        .SetInfo "Reporte", "J. Elihu", "VB6.exe"   '-- Doc Info
        .Save App.Path & "\doc1.pdf", True          '-- Save & Open
        
    End With
    Set Obj = Nothing
End Sub

Private Sub Demo2()
Dim Obj As New cCraft

    With Obj
        
        '-- Manually add the first page
        .StartDoc PAGE_SIZE_A4, AddNewPage:=False
        
        Do
            .AddPage PORTRAIT
            
            '-- Render Watermark
            .DrawForm "Watermark"
            
            '-- DRAW PAGE
            .SetFont HELVETICA_BOLD, 12
            .DrawText "Contenido página " & .PageCount, 40, 120, T_CENTER, .PageWidth - 80
            .SetFont HELVETICA, 10
            .DrawText FILE("info.txt"), 40, 150, T_JUSTIFIED, .PageWidth - 80, 8
            
            
            '-- Draw Page Number
            .SetFont HELVETICA, 10
            .DrawText "Página " & .PageCount, 40, .PageHeight - 40, T_RIGHT, .PageWidth - 150
            
            '-- Render Header & Fotter
            .DrawForm "Header"
            .DrawForm "Fotter"
            
        Loop Until Obj.PageCount = 12

        '-- Form: Watermark
        .AddForm "Watermark"
        .DrawImage "ImgWm", (.PageWidth - 150) / 2, (.PageHeight - 150) / 2, 150, 150
        
        '-- Form: Header
        .AddForm "Header", 0, 0, .PageWidth, 70
        .SetFont HELVETICA_BOLD, 14
        .DrawText "cCraft", 40, 35
        .DrawImage "logo", .PageWidth - 80, 25, 30, 30
        .DrawLine 40, 68, .PageWidth - 40, 68, &H695444, 0.5
        
        '-- Form: Fotter
        .AddForm "Fotter", 0, .PageHeight - 60, .PageWidth, 60
        .DrawLine 40, .PageHeight - 58, .PageWidth - 40, .PageHeight - 58, &HEE9800, 0.5, "3 2"
        .SetFont HELVETICA_OBLIQUE, 9, &H1F2F26
        .DrawText "cCraft - Powered By J. Elihu © 2025", 40, .PageHeight - 40
        .SetFont HELVETICA, 10, vbBlack
        .DrawText " de " & .PageCount, .PageWidth - 110, .PageHeight - 40 '/* Page Count */
        
        '-- Add Required Images
        .AddImage App.Path & "\img\32bpp_2x2.png", "logo"
        .AddImage App.Path & "\img\watermark.png", "ImgWm"
        
        
        .Save App.Path & "\doc2.pdf", True
    End With
    Set Obj = Nothing
End Sub



'----------------------------------------------------------------------------------------------------------------------
' TODO: Random data (Sample)
'======================================================================================================================
Public Function RandomName() As String
    Static vData As Variant
    If IsEmpty(vData) Then
        vData = Split("Carlos|Lucía|Mateo|Sofía|Andrés|Valentina|José|Camila|Luis|María|Jorge|Daniela|Raúl|Fernanda|Pedro|Laura|Diego|Ana|Felipe|Gabriela|Manuel|Isabel|Ricardo|Claudia|Hugo|Paula|Iván|Verónica|Óscar|Natalia", "|")
    End If
    RandomName = vData(RandomNum(UBound(vData)))
End Function
Public Function RandomLastName() As String
    Static vData As Variant
    If IsEmpty(vData) Then _
    vData = Split("García|Rodríguez|Martínez|López|Hernández|González|Pérez|Sánchez|Ramírez|Torres|Flores|Rivera|Gómez|Díaz|Cruz|Vargas|Morales|Ortiz|Silva|Ramos|Castillo|Jiménez|Reyes|Mendoza|Ruiz|Guerrero|Navarro|Romero|Delgado|Campos", "|")
    RandomLastName = vData(RandomNum(UBound(vData)))
End Function
Public Function RandomJob() As String
    Static vData As Variant
    If IsEmpty(vData) Then _
    vData = Split("Ingeniero|Doctor|Profesor|Abogado|Contador|Enfermero|Arquitecto|Electricista|Carpintero|Chef|Policía|Soldador|Diseñador|Programador|Mecánico", "|")
    RandomJob = vData(RandomNum(UBound(vData)))
End Function
Public Function RandomNum(Max As Long) As Long
    RandomNum = Int(Max * Rnd)
End Function
Private Function FILE(Fn As String) As String
Dim Out As String
Dim FF  As Integer
    FF = FreeFile
    Open App.Path & "\" & Fn For Input As #FF
    Out = Input$(LOF(FF), FF)
    Close #FF
    FILE = Out
End Function

