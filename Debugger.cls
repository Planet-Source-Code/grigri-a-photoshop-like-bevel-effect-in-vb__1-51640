VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Debugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents picOut As PictureBox
Attribute picOut.VB_VarHelpID = -1
Private m_w As Long, m_h As Long

Private m_CellX As Long, m_CellY As Long
Private m_CellW As Long, m_CellH As Long
Private m_ClientOffX As Long, m_ClientOffY As Long
Private m_rowHeight As Long
Private m_SpacingX As Long, m_SpacingY As Long

Public Property Set Target(picTarget As PictureBox)
    Set picOut = picTarget
    With picOut
        .ScaleMode = vbPixels
        m_w = .ScaleWidth
        m_h = .ScaleHeight
    End With
    Clear
End Property

Public Sub SetSpacing(ByVal dX As Long, ByVal dy As Long)
    m_SpacingX = dX
    m_SpacingY = dy
End Sub

Public Sub NewCell(ByVal CellWidth As Long, ByVal CellHeight As Long, CellTitle As String)
    ' Update X from old cell
    If m_CellW > 0 Then
        m_CellX = m_CellX + m_CellW + m_SpacingX
    End If
    
    Dim txtWidth As Long, txtHeight As Long, txtX As Long, txtY As Long
    txtWidth = picOut.TextWidth(CellTitle)
    txtHeight = picOut.TextHeight(CellTitle)
    If txtWidth > CellWidth Then
        m_CellW = txtWidth
        m_ClientOffX = (txtWidth - CellWidth) / 2
        txtX = 0
    Else
        m_CellW = CellWidth
        m_ClientOffX = 0
        txtX = (CellWidth - txtWidth) / 2
    End If
    
    ' Now we know the width of the cell, check we have enough space
    If m_CellX + m_CellW > m_w Then
        ' Not enough space, new row
        m_CellY = m_CellY + m_rowHeight + m_SpacingY
        ' Reset X to start
        m_CellX = m_SpacingX
    End If
    
    m_ClientOffY = txtHeight
    m_CellH = CellHeight + txtHeight
    If m_CellH > m_rowHeight Then m_rowHeight = m_CellH
    picOut.CurrentX = m_CellX + txtX
    picOut.CurrentY = m_CellY + txtY
    picOut.Print CellTitle
End Sub

Public Sub Plot(ByVal X As Long, ByVal Y As Long, ByVal CLR As Long)
    picOut.PSet (X + m_CellX + m_ClientOffX, Y + m_CellY + m_ClientOffY), CLR
End Sub

Public Sub Clear()
    picOut.Cls
    m_CellX = m_SpacingX
    m_CellY = m_SpacingY
    m_rowHeight = 0
    m_ClientOffX = 0
    m_ClientOffY = 0
End Sub

Private Sub picOut_Resize()
    m_w = picOut.ScaleWidth
    m_h = picOut.ScaleHeight
End Sub
