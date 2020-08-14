#Region "Imports"
Imports System.Diagnostics
Imports AmosEngineLib.AmosEngine.TMatrixID
Imports System.Xml
Imports Amos
Imports System.Collections.Generic
Imports System.Linq
#End Region

<System.ComponentModel.Composition.Export(GetType(Amos.IPlugin))>
Public Class FitClass
    Implements Amos.IPlugin
    'This plugin was written by Matthew James May 2019 for James Gaskin

    Public Function Name() As String Implements Amos.IPlugin.Name
        Return "J-N Plot Analysis"
    End Function

    Public Function Description() As String Implements Amos.IPlugin.Description
        Return "Assesses the affect of moderators on a specific factor. See statwiki.kolobkreations.com for more information."
    End Function

    Public Function Mainsub() As Integer Implements Amos.IPlugin.MainSub
        'Identification of Variables
        Dim sDependent As String = ""
        Dim sInteraction As String = ""
        Dim sIndependent As String = ""
        Dim iSelected As Integer = 0

        'User Question: Have You Selected All 3 Important Paths + Moderator. If not, exit function.
        If MsgBox("Have you selected the two paths (independent-->dependent | interaction-->dependent), the covariance (independent<-->interaction), and the interaction?", vbYesNo) = vbNo Then
            MsgBox("Please select the following: independent-->dependent | interaction-->dependent |  independent<-->interaction | interaction")
            Exit Function
        End If

        'Step 1: Identify selected paths and variables. Store names into variables.
        For Each variable In Amos.pd.PDElements 'Always Processes In This Order: Variables, Paths, Covariances
            If variable.IsSelected Then
                Select Case True
                    Case variable.IsPath 'One of two paths selected. 
                        If variable.Variable1.NameOrCaption <> sInteraction Then
                            sIndependent = variable.Variable1.NameOrCaption
                        End If
                        sDependent = variable.Variable2.NameOrCaption
                        iSelected += 1
                    Case variable.IsCovariance 'Covaried link
                        iSelected += 1
                    Case variable.IsObservedVariable AndAlso variable.IsExogenousVariable 'Interaction
                        sInteraction = variable.NameOrCaption
                        iSelected += 1
                End Select
            End If
        Next

        'At the end of loop, we should have identified four selected variables (as seen above). If something other than four selected, throw error and end function.
        If iSelected <> 4 Then
            MsgBox("You have not selected the correct amount of elements. Please try again.")
            Exit Function
        End If

        Amos.pd.AnalyzeCalculateEstimates()

        JNAnalysis(sInteraction, sIndependent, sDependent)
    End Function

    Function JNAnalysis(sInteraction As String, sIndependent As String, sDependent As String)
        '--STEP 2: CALCULATE PLOT POINTS-- (INPUTS = sInteraction, sIndependent, sDependent)
        Dim dZValue As Double = 1.96
        Dim dY1 As Double = 0
        Dim dY3 As Double = 0
        Dim dB1 As Double = 0
        Dim dB1B3 As Double = 0
        Dim dB3 As Double = 0
        Dim dSE As Double = 0
        Dim dCI As Double = 0
        Dim dSlope As Double = 0
        Dim A As Double = 0
        Dim B As Double = 0
        Dim C As Double = 0
        Dim dROS1 As Double = 0
        Dim dROS2 As Double = 0
        Dim dSlope1 As Double = 0
        Dim dSlope2 As Double = 0
        Dim dPValue As Double = 0

        Dim liLowerCI As New List(Of Double)
        Dim liUpperCI As New List(Of Double)
        Dim liSimpleSlopes As New List(Of Double)
        Dim liModerators As New List(Of Double)

        Dim xmlRegression As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Regression Weights:']/table/tbody")

        For x = 1 To GetNodeCount(xmlRegression)
            Select Case MatrixName(xmlRegression, x, 0) + MatrixName(xmlRegression, x, 2)
                Case sDependent + sInteraction
                    dPValue = MatrixElement(xmlRegression, x, 6)
                    dY3 = MatrixElement(xmlRegression, x, 3)
                    dB3 = MatrixElement(xmlRegression, x, 4) ^ 2
                Case sDependent + sIndependent
                    dY1 = MatrixElement(xmlRegression, x, 3)
                    dB1 = MatrixElement(xmlRegression, x, 4) ^ 2
            End Select
        Next

        Dim xmlCovariance As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Covariances:']/table/tbody")

        For x = 1 To GetNodeCount(xmlCovariance)
            If ((MatrixName(xmlCovariance, x, 0) + MatrixName(xmlCovariance, x, 2)) = (sInteraction + sIndependent)) Or ((MatrixName(xmlCovariance, x, 0) + MatrixName(xmlCovariance, x, 2)) = (sIndependent + sInteraction)) Then
                dB1B3 = MatrixElement(xmlCovariance, x, 3) ^ 2
                Exit For
            End If
        Next

        Dim i As Integer
        i = 0

        For dM = -1 To 1 Step 0.1
            dSlope = dY1 + (dY3 * dM)
            dSE = Math.Sqrt(dB1 + 2 * dM * dB1B3 + dM ^ 2 * dB3)
            dCI = dSE * dZValue

            'MsgBox("test")
            liLowerCI.Add(dSlope - dCI)
            liUpperCI.Add(dSlope + dCI)
            liSimpleSlopes.Add(dSlope)
            liModerators.Add(dM)
            i += 1
        Next

        A = dZValue ^ 2 * dB3 - dY3 ^ 2
        B = 2 * ((dZValue ^ 2 * dB1B3) - dY1 * dY3)
        C = dZValue ^ 2 * dB1 - dY1 ^ 2

        dROS1 = ((-B) - Math.Sqrt(B ^ 2 - 4 * A * C)) / (2 * A)
        dROS2 = ((-B) + Math.Sqrt(B ^ 2 - 4 * A * C)) / (2 * A)

        'Set up the listener To output the debugs
        If (System.IO.File.Exists("JN.html")) Then
            System.IO.File.Delete("JN.html")
        End If
        Dim debug As New AmosDebug.AmosDebug
        Dim resultWriter As New TextWriterTraceListener("JN.html")
        Trace.Listeners.Add(resultWriter)

        'Write the beginning of the document
        debug.PrintX("<html lang=""en"" dir=""ltr"">
                <head>
                  <meta charset=""utf-8"">
                  <script src=""https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.8.0/Chart.min.js""></script>
                  <script src=""https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-annotation/0.5.7/chartjs-plugin-annotation.min.js""></script>
                  <title></title>
                </head>

                <body>")

        debug.PrintX("<h1>Johnson-Neyman Plot</h1></hr>")

        debug.PrintX("<canvas id=""myChart"" width=""400"" height=""400""></canvas><br>")

        debug.PrintX("<h2>Analysis</h2>")

        debug.PrintX("<table><tr><th>Interaction</th><th>Estimate</th><th>Standard Error</th><th>P-Value</th></tr><tr><td>" + sDependent + " <-- " + sInteraction + "</td><td>" + dY3.ToString("0.000") + "</td><td>" + Math.Sqrt(dB3).ToString("0.000") + "</td><td>" + dPValue.ToString("0.000") + "</td></tr></table>")




        If Double.IsNaN(dROS1) AndAlso Double.IsNaN(dROS2) Then 'CONDITION: Both are complex numbers (NO EFFECT)
            debug.PrintX("<p>There is no meaningful moderation because Y = 0 falls within all possible X values inside the confidence intervals.</p>")
        Else
            If (dROS1 > 1 OrElse dROS1 < -1) AndAlso (dROS2 > 1 OrElse dROS2 < -1) Then 'Both Intercepts Are Out Of Range (FULL EFFECT)
                If dPValue <= 0.05 Then '(SIGNIFICANT FULL)
                    debug.PrintX("<p>Moderation is present because Y does not equal zero for any values of X within the confidence intervals within a relevant range of X values.</p>")
                Else '(NON-SIGNIFICANT FULL)
                    debug.PrintX("<p>Although there was a non-significant interaction effect (P-Value = " + dPValue.ToString("0.000") + "), we may still consider moderation to be present because Y does not equal zero for any values of X within the confidence intervals within a relevant range of X values.</p>")
                End If
            Else
                If dPValue <= 0.05 Then '(SIGNIFICANT PARTIAL BASE)
                    debug.PrintX("<p>Moderation is present because Y does not equal zero for any values of X within the confidence intervals within a relevant range of X values. ")
                Else '(NON-SIGNIFICANT PARTIAL BASE)
                    debug.PrintX("<p>Although there was a non-significant interaction effect (P-Value = " + dPValue.ToString("0.000") + "), we may still consider moderation to be present for all values of X where Y = 0 does not fall within the confidence intervals. ")
                End If
                If (dROS1 < 1 AndAlso dROS1 > -1) AndAlso (dROS2 < 1 AndAlso dROS2 > -1) Then 'Both Intercepts Are Within Range (Significant In A Certain Range)
                    dSlope1 = dY1 + (dY3 * dROS1)
                    dSlope2 = dY1 + (dY3 * dROS2)
                    If (dSlope1 * dSlope2) > 0 Then '(SIGNIFICANT BETWEEN RANGE)
                        debug.PrintX("In this particular interaction, Y = 0 does not fall within the confidence interval for all values of X greater than " + dROS1.ToString("0.000") + " and less than " + dROS2.ToString("0.000") + ".</p")
                    Else '(SIGNIFICANT ABOVE/BELOW RANGE)
                        debug.PrintX("In this particular interaction, Y = 0 does not fall within the confidence interval for all values of X less than " + dROS1.ToString("0.000") + " or greater than " + dROS2.ToString("0.000") + ".</p")
                    End If
                ElseIf (dROS1 < 1 AndAlso dROS1 > -1) Then 'ROS1 Within Range
                    dSlope1 = dY1 + (dY3 * dROS1)
                    If dSlope1 > 0 Then '(SIGNIFICANT ABOVE RANGE)
                        debug.PrintX("In this particular interaction, Y = 0 does not fall within the confidence interval for all values of X greater than " + dROS1.ToString("0.000") + ".</p")
                    Else '(SIGNIFICANT BELOW RANGE)
                        debug.PrintX("<p>In this particular interaction, Y = 0 does not fall within the confidence interval for all values of X less than " + dROS1.ToString("0.000") + ".</p")
                    End If
                ElseIf (dROS2 > 1 OrElse dROS2 < -1) Then 'ROS2 Within Range
                    dSlope2 = dY1 + (dY3 * dROS2)
                ElseIf dSlope1 > 0 Then '(SIGNIFICANT ABOVE RANGE)
                    debug.PrintX("In this particular interaction, Y = 0 does not fall within the confidence interval for all values of X greater than " + dROS2.ToString("0.000") + ".</p")
                Else '(SIGNIFICANT BELOW RANGE)
                    debug.PrintX("In this particular interaction, Y = 0 does not fall within the confidence interval for all values of X less than " + dROS2.ToString("0.000") + ".</p")
                End If

            End If
        End If

        debug.PrintX("<h3>References</h3><hr><p>--Hayes, A. F., and Matthes, J. 2009. ""Computational Procedures for Probing Interactions in OLS and Logistic Regression: SPSS and SAS Implementations,"" Behavior research methods (41:3), pp. 924-936.</p>")
        debug.PrintX("<p>--Gaskin, J. and James, M. (2019) “"Johnson-Neyman Plot Analysis Plugin for AMOS”". <a href=""http://statwiki.kolobkreations.com"">Gaskination's StatWiki</a>.</p>")

        debug.PrintX("<script>
    var lowerInt = " + Str(dROS1) + ";
    var upperInt = " + Str(dROS2) + ";
    var xTitle = """ + sIndependent + """;
    var yTitle = """ + sDependent + """;
    var mTitle = ""Moderator"";
    var xPoints = [")

        For x = 0 To liModerators.Count - 1
            If x <> liModerators.Count - 1 Then
                debug.PrintX(Str(liModerators(x)) + ", ")
            Else
                debug.PrintX(Str(liModerators(x)) + "]; var slopePoints = [")
            End If
        Next
        For x = 0 To liSimpleSlopes.Count - 1
            If x <> liSimpleSlopes.Count - 1 Then
                debug.PrintX(Str(liSimpleSlopes(x)) + ", ")
            Else
                debug.PrintX(Str(liSimpleSlopes(x)) + "]; var uCIPoints = [")
            End If
        Next
        For x = 0 To liUpperCI.Count - 1
            If x <> liUpperCI.Count - 1 Then
                debug.PrintX(Str(liUpperCI(x)) + ", ")
            Else
                debug.PrintX(Str(liUpperCI(x)) + "]; var lCIPoints = [")
            End If
        Next

        For x = 0 To liLowerCI.Count - 1
            If x <> liLowerCI.Count - 1 Then
                debug.PrintX(Str(liLowerCI(x)) + ", ")
            Else
                debug.PrintX(Str(liLowerCI(x)) + "];")
            End If
        Next

        debug.PrintX("
                var storage1 = [];
                var storage2 = [];
                var storage3 = [];
                for (var i = 0; i < xPoints.length; i++) {
                  x = xPoints[i];
                  slopeY = slopePoints[i];
                  uCIY = uCIPoints[i];
                  lCIY = lCIPoints[i];
                  var jsonSlope = {
                    x: x,
                    y: slopeY
                  };
                  var jsonUCI = {
                    x: x,
                    y: uCIY
                  };
                  var jsonLCI = {
                    x: x,
                    y: lCIY
                  };
                  storage1.push(jsonSlope);
                  storage2.push(jsonUCI);
                  storage3.push(jsonLCI);
                }
                var ctx = document.getElementById('myChart');
                var myChart = new Chart(ctx, {
                  type: 'line',
                  data: {
                    labels: [],
                    datasets: [{
                        label: 'Simple Slope',
                        data: storage1,
                        borderColor: 'rgb((125, 125, 125))',
                        backgroundColor: 'rgb(232, 233, 235)',
                        borderDash: [10, 3],
                        borderWidth: 1,
                        fill: false,
                        showLine: true
                      },
                      {
                        label: 'Upper CL',
                        data: storage2,
                        borderColor: 'rbg(125, 125, 125)',
                        borderWidth: 1.5,
                        backgroundColor: 'rgba(232, 233, 235, 0.7)',
                        fill: false,
                        showLine: true
                      },
                      {
                        label: ""Lower CI"",
                        data: storage3,
                        borderColor: 'rgb(125, 125, 125)',
                        backgroundColor: 'rgba(232, 233, 235, 0.7)',
                        borderWidth: 1.5,
                        fill: '-1',
                        showLine: true
                      }
                    ],
                  },
                  options: {
                    legend: {
                      display: false,
                      position: ""right""
                    },
                    scales: {
                      yAxes: [{
                        scaleLabel: {
                          display: true,
                          labelString: ""Simple Slope Of "" + xTitle + "" (X) predicting "" + yTitle + "" (Y)"",
                          fontStyle: ""bold""
                        },
                        ticks: {
                          beginAtZero: false
                        },
                        gridLines: {
                          color: ""rgba(0, 0, 0, 0)"",
                        }
                      }],
                      xAxes: [{
                        type: 'linear',
                        scaleLabel: {
                          display: true,
                          labelString: ""Value Of "" + mTitle + "" (M)"",
                          fontStyle: ""bold""
                        },
                        ticks: {
                          maxTicksLimit: 13,
                          stepSize: 0.1,
                          min: -1,
                          max: 1
                        },
                        gridLines: {
                          color: ""rgba(0, 0, 0, 0)"",

                        }
                      }]
                    },
                    elements: {
                      point: {
                        radius: 0
                      }
                    },
                    annotation: {
                      annotations: [{
                        id: 'lower',
                        type: 'line',
                        mode: 'vertical',
                        scaleID: 'x-axis-0',
                        value: lowerInt,
                        borderColor: 'rgb(125, 125, 125)',
                        borderWidth: 1.5
                      },
                      {
                        id: 'upper',
                        type: 'line',
                        mode: 'vertical',
                        scaleID: 'x-axis-0',
                        value: upperInt,
                        borderColor: 'rgb(125, 125, 125)',
                        borderWidth: 1.5
                      }]
                    }
                  }
                });
              </script>
            <style>
                canvas {
                  width: 100% !important;
                  max-width: 800px;
                  height: auto !important; }
                table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}
              </style>
    
            </body>

            </html>
            ")

        'Take down our debugging, release file, open html
        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Process.Start("JN.html")
    End Function

    'Finds number of nodes in an xml table
    Function GetNodeCount(table As XmlElement) As Integer

        Dim nodeCount As Integer = 0

        'Handles a model with zero correlations
        Try
            nodeCount = table.ChildNodes.Count
        Catch ex As NullReferenceException
            nodeCount = 0
        End Try

        GetNodeCount = nodeCount

    End Function

    'Use an output table path to get the xml version of the table.
    Public Function GetXML(path As String) As XmlElement

        'Gets the xpath expression for an output table.
        Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
        doc.Load(Amos.pd.ProjectName & ".AmosOutput")
        Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
        Dim eRoot As Xml.XmlElement = doc.DocumentElement

        Return eRoot.SelectSingleNode(path, nsmgr)

    End Function

    'Get a string element from an xml table.
    Function MatrixName(eTableBody As XmlElement, row As Long, column As Long) As String

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixName = e.InnerText
        Catch ex As Exception
            MatrixName = ""
        End Try

    End Function

    'Get a number from an xml table
    Function MatrixElement(eTableBody As XmlElement, row As Long, column As Long) As Double

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixElement = CDbl(e.GetAttribute("x"))
        Catch ex As Exception
            MatrixElement = 0
        End Try

    End Function


End Class