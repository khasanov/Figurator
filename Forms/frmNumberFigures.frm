VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNumberFigures 
   Caption         =   "Нумерация рисунков"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   OleObjectBlob   =   "frmNumberFigures.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNumberFigures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public nUserUnit As Integer ' Единицы измерения, установленые пользователем
Public dXOrigin, dYOrigin As Double ' Начало отсчета, установленное пользователем
Public dXCoord, dYCoord As Double ' Координаты подписей
Public dFontSize As Double ' Размер шрифта
Public bFontItalic As cdrTriState ' Курсив
Public bFontBold As cdrTriState ' Жирный

Dim collPages As Pages
Dim collShapes As Shapes
Dim sLabel As Shape

' Удаление меток
Public Sub DeleteLabels(ByVal strName As String)

Dim s As Shape
Dim p As Page

For Each p In collPages
  Set collShapes = p.Shapes
  For Each s In collShapes
    If s.Name = strName Then
      s.Delete
    End If
  Next s
Next p

End Sub

Private Sub cboFontSize_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

  If Not ((Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Or (Chr(KeyAscii) = ".")) Then
    KeyAscii = 0
  End If

End Sub

' Нажата кнопка "Применить"
Private Sub cmdApply_Click()

  'This script will add page numbering to all pages in a document
   
  'Object Variables
  Dim p As Page
  Dim l As Layer
    
  'Variable Declaration
  Dim nNumOfPages, nCurrentPageNumber As Integer
  Dim nCount As Integer
  Dim strLabelName, strLabelText As String
  Dim dPageWidth, dPageHeight, dShapeWidth, dShapeHeight As Double
  Dim StartPage As Integer
  
  ' Инициализация переменных
  StartPage = Int(Val(txtStartPage.Text))
  
  If cboFontSize.Text = "" Then
    dFontSize = 12
  Else
    dFontSize = Val(cboFontSize.Text)
  End If
  
  bFontBold = boldCheckBox.Value
  bFontItalic = italicCheckBox.Value
   
  If xBoxEdit.Text = "" Then
    dXCoord = 55
  Else
    dXCoord = Val(xBoxEdit.Text)
  End If
  
  If yBoxEdit.Text = "" Then
    dYCoord = 55
  Else
    dYCoord = Val(yBoxEdit.Text)
  End If
  
  strLabelName = "lblFigure"
    
  ' Общее число страниц
  nNumOfPages = collPages.Count
  
  ' Удаление меток
  Call DeleteLabels(strLabelName)
  
  ' Добавление меток
  For Each p In collPages
  
    p.Activate
    Set l = p.ActiveLayer
    
    'First we must check for an empty string
    'If we find on, we default to #
    If txtLabel.Text = "" Then
      strLabelText = "#"
    Else
      strLabelText = txtLabel.Text
    End If
    ' Замена '#' текущим номером
    strLabelText = Replace(strLabelText, "#", CStr(p.Index + StartPage - 1))
      
    ' Создание нового shape object (метки) и расположение в (0,0)
    Set sLabel = l.CreateArtisticText(0, 0, strLabelText, Font:=cboFontFace.Text, Size:=dFontSize, Bold:=bFontBold, Italic:=bFontItalic)
    
    sLabel.Name = strLabelName ' Object Name
               
    'Get Height and Width of Both the Active Page and the Label
    'These values will be used in calculating the x, y position
    'for the label. It has been done this way to ensure the label
    'will be placed approprielty on any page orientation or size.
    dShapeWidth = sLabel.SizeWidth
    dShapeHeight = sLabel.SizeHeight
    dPageWidth = p.SizeWidth
    dPageHeight = p.SizeHeight
    
    'Will run calculations to set the correct x, y values
    'Call SetLabelPosition(dPageWidth, dPageHeight, dShapeWidth, dShapeHeight, cboAlign.Text)
    
    ' Расположение метки
    sLabel.PositionX = dXCoord
    sLabel.PositionY = dYCoord

  Next p
  
End Sub

' Нажата кнопка "Очистить"
Private Sub cmdClear_Click()

  Call DeleteLabels("lblFigure")

End Sub

' Нажата кнопка "Выйти"
Private Sub cmdExit_Click()

  Unload Me

End Sub

' Нажат чекбокс "курсив"
Private Sub italicCheckBox_Click()

  bFontItalic = italicCheckBox.Value
  
End Sub

' Нажат чекбокс "жирный"
Private Sub boldCheckBox_Click()

  bFontBold = boldCheckBox.Value
  
End Sub

' Изменено положение метки
Private Sub xBoxEdit_Change()

  If xBoxEdit.Value = "" Then
    xBoxEdit.Value = 0
  End If
    
  dXCoord = Val(xBoxEdit.Text)
  
End Sub

' Изменено положение метки
Private Sub yBoxEdit_Change()

  If yBoxEdit.Value = "" Then
    yBoxEdit.Value = 0
  End If
    
  dYCoord = Val(yBoxEdit.Text)
  
End Sub


Private Sub txtStartPage_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  
  If txtStartPage.Value = "" Then
    txtStartPage.Value = 1#
  End If
End Sub

Private Sub txtStartPage_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

  If Not (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Then
    KeyAscii = 0
  End If

End Sub

Private Sub UserForm_Initialize()

  Dim nCount, nNumOfFonts As Integer
  Dim nDefaultFont As Integer
  Dim strDefaultFont As String
  Dim d As Document, v As Variant
   
  On Error GoTo mainErrHandler
  
  Set d = ActiveDocument
  nNumOfFonts = FontList.Count
  strDefaultFont = "Times New Roman"
  nDefaultFont = 1
  
  'Let's first initialize all our collections
  Set collPages = ActiveDocument.Pages
  
  'The FontList object contains all the font the user can use
  'Let's put these font face names into a list box
  nCount = 1
  For Each v In FontList
    cboFontFace.AddItem v
    'Let's find a default font to set
    If v = strDefaultFont Then
      'Here we have found the default font strDefaultFont
      nDefaultFont = nCount
    End If
    nCount = nCount + 1
  Next v
  
  'Now we set the default font, if one is not found, we then
  'default to the first item in the list
  cboFontFace.ListIndex = nDefaultFont - 1
  
  '...and now, we'll dynamically add list values for the Font Size drop down
  cboFontSize.AddItem 6
  cboFontSize.AddItem 7
  cboFontSize.AddItem 8
  cboFontSize.AddItem 9
  cboFontSize.AddItem 10
  cboFontSize.AddItem 11
  cboFontSize.AddItem 12
  cboFontSize.AddItem 14
  cboFontSize.AddItem 16
  cboFontSize.AddItem 18
  cboFontSize.AddItem 24
  cboFontSize.AddItem 36
  cboFontSize.AddItem 48
  cboFontSize.AddItem 72
  cboFontSize.AddItem 100
  cboFontSize.AddItem 150
  cboFontSize.AddItem 200
  cboFontSize.ListIndex = 6
  
  ' Переходим к нашей системе коодринат
  ' Сначала сохраняем настройки пользователя
  nUserUnit = d.Unit
  dXOrigin = d.DrawingOriginX
  dYOrigin = d.DrawingOriginY
    
  '...затем устанавливаем наши
  d.Unit = cdrMillimeter
  d.DrawingOriginX = -105 '-4.25
  d.DrawingOriginY = -148 '-5.5
    
mainErrHandler:

If Err.Number > 0 Then
  MsgBox Err.Number & " " & Err.Description
  End
End If

End Sub

Private Sub UserForm_Terminate()

  Dim d As Document
  Set d = ActiveDocument
  ' Восстанавливаем настройки пользователя
  d.Unit = nUserUnit
  d.DrawingOriginX = dXOrigin
  d.DrawingOriginY = dYOrigin

End Sub
