Attribute VB_Name = "modRoundText"
'------------------------------------------------------------
' Nome del Progetto: Project1
' Nome del Modulo: modRoundText
' Scopo:
' Data: 07/05/2001
' Ora: 12.29
' Revisione:
' Autore: NDV Software
'------------------------------------------------------------
' ****************************************************************************************************
' Copyright © 1990 - 2001 NDV Software,
' Tutti i diritti riservati, ndv@interfree.it
' ****************************************************************************************************
Option Explicit
Global Const PIGRECO = 3.141592654

Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFacename As String * 33
End Type

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'------------------------------------------------------------
' Nome: drawCircularText
' Scopo: Disegna un testo su cerchio o settore circolare
' Visibilità: Public
' Parametri:
'    Obj As Object              Oggetto su cui stampare Picture o Printer
'    Testo As String            Testo da stampare
'    TextStartAngle As Single   Angolo di inizio stampa del testo
'    Raggio As Single           Raggio del cerchio o del settore circolare
'    CX As Integer              Coordinata X del Centro del cerchio
'    CY As Integer              Coordinata Y del Centro del cerchio
'    TextSector As Single       Porzione di settore circolare da coprire con il testo
'                               180 -> Metà cerchio  360 -> tutto il cerchio ecc...
'    Tutti i parametri grafici quali Font Name,colore, dimensione del font, bold, italic
'    vanno settati direttamente sull'oggetto output della stampa
' Data: lunedì 7 maggio 2001
' Ora: 12.25
' Autore: NDV Software
' Revisione:
'------------------------------------------------------------
Public Sub drawCircularText(Obj As Object, Testo As String, TextStartAngle As Single, Raggio As Single, CX As Integer, CY As Integer, TextSector As Single)
  On Error GoTo Errore
  Dim F As LOGFONT
  Dim hPrevFont As Long
  Dim hFont As Long
  Dim FontName As String
  Dim I As Integer
  Dim Passo As Single
  
  Passo = TextSector / Len(Testo)   'Angular Step
    
  For I = 1 To Len(Testo)
    F.lfEscapement = 10 * TextStartAngle - (10 * Passo * (I - 1)) 'rotation angle, in tenths (x10)
    FontName = Obj.FontName + Chr$(0) 'null terminated
    F.lfFacename = FontName
    F.lfHeight = (Obj.FontSize * -20) / Screen.TwipsPerPixelY
    hFont = CreateFontIndirect(F)
    hPrevFont = SelectObject(Obj.hdc, hFont)
    Obj.CurrentX = CX + Raggio * Sin((-180 + TextStartAngle - (I - 1) * Passo) * PIGRECO / 180)
    Obj.CurrentY = CY + Raggio * Cos((-180 + TextStartAngle - (I - 1) * Passo) * PIGRECO / 180)
    'Obj.Line (CX, CY)-(CurrentX, CurrentY)
    Obj.Print Mid(Testo, I, 1)
    hFont = SelectObject(Obj.hdc, hPrevFont)
    DeleteObject hFont
  Next I
  
  Exit Sub
Errore:
  Exit Sub
End Sub
