Attribute VB_Name = "UsefulFuncs"
'Copyright © 2020 Ben Randerson

Option Explicit
'@Module: This module contains a set of useful functions


Public Function LINEAR_INTERP(ByVal x As Double, xValues As Range, yValues As Range) As Double

    '@Description: This function performs linear interpolation
    '@Author: Ben Randerson
    '@Version: 1.0.0
    '@Param: x is x value to be interpolated
    '@Param: xValues is the range of x-values
    '@Param: yValues is the range of y-values
    '@Returns: Returns the interpolated value
    '@Warning: x-values must be ascending from top to bottom
    '@Warning: y-values must be ascending from top to bottom
    '@Warning: x and y must be within the range of available data
    '@Example: =LINEAR_INTERP(1,A1:A4,B1:B4) -> 2.5

    Dim x1 As Double
    Dim x2 As Double
    Dim y1 As Double
    Dim y2 As Double
    
    x1 = Application.WorksheetFunction.Index(xValues, Application.WorksheetFunction.Match(x, xValues, 1))
    x2 = Application.WorksheetFunction.Index(xValues, Application.WorksheetFunction.Match(x, xValues, 1) + 1)
    y1 = Application.WorksheetFunction.Index(yValues, Application.WorksheetFunction.Match(x, xValues, 1))
    y2 = Application.WorksheetFunction.Index(yValues, Application.WorksheetFunction.Match(x, xValues, 1) + 1)
    
    LINEAR_INTERP = y1 + (y2 - y1) * (x - x1) / (x2 - x1)
    
End Function

