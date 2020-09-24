VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************
' Name: stack
' Description:Defines a class "stack" wi
'     th the usual operations of pop and push
' By: Manuel Jose Domingues Alves
'
' Inputs:A variant to the "push" method
'
' Returns:A variant from the "pop" metho
'     d; a property "stackLevel" that states t
'     he number of elements in the stack
'
' Assumes:VB5
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/xq/ASP/txtCode
'     Id.882/lngWId.1/qx/vb/scripts/ShowCode.h
'     tm'for details.'**************************************

' This code shows a way on how to
' create and use simple stacks in VB5
'
' 1 - create a new "standard EXE" projec
'     t
' 2 - go to the menu "project", select
' "add class module" and choose
' "classe module"
' 3 - in the "properties window" change
' its name to "stack"
' 4 - paste the following code:
'
Option Explicit
Private stackValue() As Variant
Private level As Long


Public Property Get stackLevel() As Integer
    stackLevel = UBound(stackValue)
End Property


Private Sub Class_Initialize()
    ReDim stackValue(0)
    level = 0
End Sub

Public Function Clear() As Variant
    ReDim stackValue(0)
    level = 0
End Function

Public Sub push(elem As Variant)
    level = level + 1
    ReDim Preserve stackValue(level)
    stackValue(level) = elem
End Sub

Public Function pop() As Variant
    Dim res As Variant


    If level > 0 Then
        res = stackValue(level)
        level = level - 1
        ReDim Preserve stackValue(level)
    Else
        res = Null
    End If
    pop = res
End Function

Public Function popNo(WhatLevel As Long) As Variant
    Dim res As Variant


    If level > 0 And WhatLevel <= level Then
        res = stackValue(WhatLevel)
    Else
        res = Null
    End If
    popNo = res
    
End Function

