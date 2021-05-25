Attribute VB_Name = "Enums"
'@Folder("Model")
Option Explicit

Public Enum LogLevel
  aOff = 0
  bTrace = 1
  cDebug = 2
  dInfo = 3
  eWarn = 4
  fError = 5
End Enum

Public Enum RenderingColor
  NoColor = 0
  Black = 0
  Blue = 16711680
  Cyan = 16776960
  Green = 65280
  Magenta = 16711935
  Red = 255
  White = 16777215
  Yellow = 65535
End Enum

Public Enum RenderingGrouping
  IsGroup = -1
  NoGroup = 0
End Enum

Public Enum RenderingReadOnly
  RO = -1
  RW = 0
End Enum

Public Enum RenderingVisible
  Visible = -1
  Hidden = 0
End Enum

Public Enum FormattingType
  NoFormatting
  With_Formula
  Col_Op_OtherCol
  Col_Op_Today
  Col_Op_Integer
  Col_BarChart
End Enum

Public Enum FormattingOperator
  NoOperator = 0
  Equal = 3
  Greater = 5
  Less = 6
  Greater_Equal = 7
  Less_Equal = 8
End Enum

Public Enum ValidationType
  NoValidation
  Numbers
  SingleChars
  Dates
  DropDown
End Enum

