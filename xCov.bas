Sub xCov()

'' this macro calculates xcov for two dataset
'
' xcov is cross correlation of two average-corrected dataset as in MATLAB
' reference:
' https://en.wikipedia.org/wiki/Local_regression
' https://mathworks.com/help/signal/ref/xcov.html

Dim totalRange, tempRange As Range
Dim totX, cntX, cntY, sumCnt As Integer
Dim tempAve, tempSum, alphaRatio, distBase, wSum, maxVal As Double
Dim resolutionN, cntR, binNum, binEnd, cntBin As Integer
Dim adrMax As String

cntX = 1

'modify these for LOWESS: 100 is 0.01 precision (10 is 0.1)
'  alpha ratio is lowess sampling ratio, n / N = alpha where n is sample size, N is total number of one dataset
resolutionN = 100
alphaRatio = 0.25

'count data size
Do While Not IsEmpty(Cells(cntX, 1))
     cntX = cntX + 1
     Loop
     
totX = cntX - 1

' calculate bins for LOWESS
binNum = totX * alphaRatio
binEnd = WorksheetFunction.Floor(binNum / 2, 1)

' display the sum, number and average of each dataset on the last three rows
Cells(totX + 2, 1) = "sum, counts, average"
Cells(totX + 1, 3) = "average adjusted"

For cntY = 1 To 2

     Set tempRange = Range(Cells(1, cntY), Cells(totX, cntY))
     Cells(totX + 3, cntY) = WorksheetFunction.Sum(tempRange)
     Cells(totX + 4, cntY) = WorksheetFunction.Count(tempRange)
     Cells(totX + 5, cntY) = WorksheetFunction.Sum(tempRange) / WorksheetFunction.Count(tempRange)
     
     ' write average-adjusted dataset
     For cntX = 1 To totX
          Cells(cntX, cntY + 2) = Cells(cntX, cntY) - Cells(totX + 5, cntY)
          Next cntX
     
     Next cntY

' add spacers in first dataset for easy calculation
For cntX = 1 To totX - 1
     Cells(cntX, 5) = 0
     Next cntX
     
For cntX = totX To totX * 2 - 1
     Cells(cntX, 5) = Cells(cntX - totX + 1, 3)
     Next cntX
     
For cntX = totX * 2 To totX * 3 - 1
     Cells(cntX, 5) = 0
     Next cntX

' label timeline
For cntX = 1 To totX * 2 - 1
     Cells(cntX, 6) = cntX - totX
     Next cntX
     
'' calculate xcov
For cntX = 1 To totX * 2 - 1
     tempSum = 0
     For sumCnt = 1 To totX
          tempSum = tempSum + Cells(cntX + sumCnt - 1, 5) * Cells(sumCnt, 4)
          Next sumCnt
     Cells(cntX, 7) = tempSum
     Next cntX
     
'' calculate LOWESS
For cntX = binEnd To totX * 2 - binEnd - 2
     For cntR = 0 To resolutionN - 1
          Cells((cntX - binEnd) * 10 + cntR + 1, 9) = Cells(cntX, 6) + 0.01 * cntR
          
          tempSum = 0
          wSum = 0
          
          distBase = Abs(Cells((cntX - binEnd) * 10 + cntR + 1, 9) - Cells(cntX - binEnd + 1, 6))
          
          'this is the actual LOWESS calculation, with tri-cube weight function
          For cntBin = 1 To binNum
               tempSum = tempSum + Cells(cntX - binEnd + cntBin, 7) * (1 - (Abs(Cells((cntX - binEnd) * 10 + cntR + 1, 9) - Cells(cntX - binEnd + cntBin, 6)) / distBase) ^ 3) ^ 3
               wSum = wSum + (1 - (Abs(Cells((cntX - binEnd) * 10 + cntR + 1, 9) - Cells(cntX - binEnd + cntBin, 6)) / distBase) ^ 3) ^ 3
               Next cntBin
               
          Cells((cntX - binEnd) * 10 + cntR + 1, 10) = tempSum / wSum 'use average weighted sum
          
          Next cntR
     Next cntX
     
End Sub
