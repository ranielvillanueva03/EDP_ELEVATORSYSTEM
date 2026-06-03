' ============================================================
'  ElevatorSystem.vb  —  4-Floor Elevator Simulator
'  Visual Basic .NET  |  Windows Forms
'
'  HOW TO RUN:
'    1. Open Visual Studio 2019 or later
'    2. File > New Project > Windows Forms App (.NET Framework) — VB.NET
'    3. Replace Form1.vb contents with this entire file
'    4. Press F5 to run
' ============================================================

Public Class Form1

    ' ── Constants ──────────────────────────────────────────
    Private Const TOTAL_FLOORS As Integer = 4
    Private Const FLOOR_HEIGHT As Integer = 140    ' pixels per floor in shaft (increased for larger UI)
    Private Const SHAFT_WIDTH As Integer = 120
    Private Const CAR_WIDTH As Integer = 80
    Private Const CAR_HEIGHT As Integer = 100

    ' Floor names
    Private ReadOnly FloorNames As String() = {"GRD", "LEV", "LEV", "TOP"}

    ' ── Elevator State ─────────────────────────────────────
    ' Start at ground floor by default
    Private currentFloor As Integer = 1
    Private targetQueue As New Queue(Of Integer)
    Private isMoving As Boolean = False
    Private isEmergency As Boolean = False
    Private doorsOpen As Boolean = False
    Private carY As Integer
    Private carTargetY As Integer
    Private doorOpenPct As Single = 0.0F
    Private doorPhase As Integer = 0

    ' ── Passenger State ────────────────────────────────────
    Private passengerFloor As Integer = 1
    Private passengerInElevator As Boolean = False
    Private passengerDest As Integer = -1

    ' ── Timers ─────────────────────────────────────────────
    Private WithEvents tmrMove As New Timer() With {.Interval = 16}
    Private WithEvents tmrDoor As New Timer() With {.Interval = 600}
    Private WithEvents tmrClock As New Timer() With {.Interval = 1000}

    ' ── UI Controls ────────────────────────────────────────
    Private WithEvents shaftPanel As Panel
    Private WithEvents logBox As ListBox
    Private lblFloorVal As Label
    Private lblDirectionVal As Label
    Private lblStatusVal As Label
    Private lblPassengerVal As Label
    Private lblClock As Label
    Private lblBigFloor As Label
    Private WithEvents btnEmergency As Button

    Private WithEvents btnSpawnPassenger As Button
    Private WithEvents btnBoardPassenger As Button
    Private cmbSpawnFloor As ComboBox
    Private cmbPassengerDest As ComboBox

    Private upButtons(TOTAL_FLOORS) As Button
    Private downButtons(TOTAL_FLOORS) As Button
    Private callButtons(TOTAL_FLOORS) As Button

    ' Dot indicators per floor (2 dots each)
    Private dotPanels(TOTAL_FLOORS) As Panel
    Private personPanels(TOTAL_FLOORS) As Panel

    ' ──────────────────────────────────────────────────────
    '  Constructor
    ' ──────────────────────────────────────────────────────
    Public Sub New()
        InitializeComponent()
        Me.Text = "ElevatorSim v1.0 — 4 Floor Elevator Control System"
        Me.Size = New Size(1100, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(18, 24, 42)
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Font = New Font("Courier New", 11)

        BuildUI()

        ' Ensure person panels reflect initial passenger state
        For i As Integer = 1 To TOTAL_FLOORS
            If personPanels(i) IsNot Nothing Then personPanels(i).Visible = False
        Next

        carY = FloorToPixelY(currentFloor)
        carTargetY = carY

        tmrClock.Start()
        AddLog("System initialised. Elevator at floor " & currentFloor)
        AddLog("Passenger waiting at floor " & passengerFloor)
        RefreshStatus()
        shaftPanel.Invalidate()
    End Sub

    ' ── Manual passenger controls handlers ──────────────
    Private Sub btnSpawnPassenger_Click(sender As Object, e As EventArgs)
        If cmbSpawnFloor Is Nothing Then Return
        Dim f As Integer = CInt(cmbSpawnFloor.SelectedItem)
        passengerFloor = f
        passengerInElevator = False
        passengerDest = -1
        AddLog("Manual: Passenger spawned at F" & passengerFloor)
        ' update person panels
        For i As Integer = 1 To TOTAL_FLOORS
            If personPanels(i) IsNot Nothing Then personPanels(i).Visible = (i = passengerFloor)
        Next
        RefreshStatus()
    End Sub

    Private Sub btnBoardPassenger_Click(sender As Object, e As EventArgs)
        ' If doors are open at current floor and passenger waiting there, board or exit
        If Not doorsOpen Then
            AddLog("Doors are closed — open doors to board/exit")
            Return
        End If

        If passengerInElevator Then
            ' Try to exit at current floor
            If passengerDest = currentFloor Then
                passengerInElevator = False
                AddLog("Manual: Passenger exited at F" & currentFloor)
                passengerFloor = currentFloor
                passengerDest = -1
                ' show person outside
                If personPanels(currentFloor) IsNot Nothing Then personPanels(currentFloor).Visible = True
            Else
                AddLog("Manual: Passenger cannot exit here; destination F" & passengerDest)
            End If
        Else
            ' Board if passenger is waiting at this floor
            If passengerFloor = currentFloor Then
                Dim dest As Integer = CInt(cmbPassengerDest.SelectedItem)
                If dest = passengerFloor Then
                    AddLog("Manual: Passenger destination must differ from origin")
                    Return
                End If
                passengerInElevator = True
                passengerDest = dest
                AddLog("Manual: Passenger boarded at F" & currentFloor & " → F" & passengerDest)
                ' Queue travel to destination
                QueueFloor(passengerDest)
                ' hide person outside
                If personPanels(currentFloor) IsNot Nothing Then personPanels(currentFloor).Visible = False
                ' mark passenger as no longer waiting outside
                passengerFloor = -1
            Else
                AddLog("No passenger waiting at this floor to board")
            End If
        End If
        RefreshStatus()
    End Sub

    Private Sub PersonPanel_Paint(sender As Object, e As PaintEventArgs)
        Dim pnl As Panel = DirectCast(sender, Panel)
        Dim g As Graphics = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        g.Clear(Color.Transparent)
        Dim floor As Integer = CInt(pnl.Tag)
        If Not (Not passengerInElevator AndAlso passengerFloor = floor) Then Return
        Dim centerX As Integer = pnl.Width \ 2
        Dim bottomY As Integer = pnl.Height - 4
        Dim personH As Integer = pnl.Height - 6
        DrawPerson(g, centerX, bottomY, personH, Color.FromArgb(220, 180, 60))
    End Sub

    ' ── Floor <-> Pixel ────────────────────────────────────
    Private Function FloorToPixelY(floor As Integer) As Integer
        Dim shaftH As Integer = TOTAL_FLOORS * FLOOR_HEIGHT
        Dim centrY As Integer = shaftH - (floor - 1) * FLOOR_HEIGHT - FLOOR_HEIGHT \ 2
        Return centrY - CAR_HEIGHT \ 2
    End Function

    Private Function PixelYToFloor(y As Integer) As Integer
        Dim shaftH As Integer = TOTAL_FLOORS * FLOOR_HEIGHT
        Dim centreY As Integer = y + CAR_HEIGHT \ 2
        ' Compute the 1-based index from the top (1 = topmost band).
        Dim indexFromTop As Integer = CInt(Math.Ceiling(centreY / CDbl(FLOOR_HEIGHT)))
        Dim f As Integer = TOTAL_FLOORS - (indexFromTop - 1)
        Return Math.Max(1, Math.Min(TOTAL_FLOORS, f))
    End Function

    ' ── Queue & Movement ───────────────────────────────────
    Private Sub QueueFloor(floor As Integer)
        If Not targetQueue.Contains(floor) Then targetQueue.Enqueue(floor)
        ' Do not start moving while doors are open (wait for doors to finish)
        If Not isMoving AndAlso Not isEmergency AndAlso Not doorsOpen Then StartMoving()
    End Sub

    Private Sub StartMoving()
        If targetQueue.Count = 0 OrElse isEmergency Then
            isMoving = False : RefreshStatus() : Return
        End If
        Dim nxt As Integer = targetQueue.Peek()
        If nxt = currentFloor Then
            targetQueue.Dequeue() : ArrivedAtFloor() : Return
        End If
        isMoving = True
        carTargetY = FloorToPixelY(nxt)
        tmrMove.Start()
        AddLog("Moving " & If(nxt > currentFloor, "UP", "DOWN") & " to Floor " & nxt)
        RefreshStatus()
    End Sub

    Private Sub tmrMove_Tick(sender As Object, e As EventArgs) Handles tmrMove.Tick
        If isEmergency Then tmrMove.Stop() : Return
        Dim speed As Integer = 3
        Dim shaftH As Integer = TOTAL_FLOORS * FLOOR_HEIGHT
        ' Clamp carTargetY and carY to valid range to avoid visual glitches
        carTargetY = Math.Max(0, Math.Min(shaftH - CAR_HEIGHT, carTargetY))
        carY = Math.Max(0, Math.Min(shaftH - CAR_HEIGHT, carY))

        If Math.Abs(carY - carTargetY) <= speed Then
            carY = carTargetY
            tmrMove.Stop()
            currentFloor = PixelYToFloor(carY)
            If targetQueue.Count > 0 Then targetQueue.Dequeue()
            ArrivedAtFloor()
        Else
            carY += If(carTargetY > carY, speed, -speed)
        End If
        shaftPanel.Invalidate()
    End Sub

    Private Sub ArrivedAtFloor()
        AddLog("Arrived at floor " & currentFloor)
        ' Manual boarding: passenger will board/exit only when user presses the boarding button
        ' Automatic boarding/exiting logic removed to allow manual control.
        ClearFloorButtons(currentFloor)
        OpenDoors()
    End Sub

    ' ── Doors ──────────────────────────────────────────────
    Private Sub OpenDoors()
        doorsOpen = True
        doorPhase = 1
        doorOpenPct = 0
        tmrDoor.Interval = 120
        tmrDoor.Start()
        AddLog("Doors opening at F" & currentFloor)
    End Sub

    Private Sub tmrDoor_Tick(sender As Object, e As EventArgs) Handles tmrDoor.Tick
        Select Case doorPhase
            Case 1
                doorOpenPct = Math.Min(1.0F, doorOpenPct + 0.2F)
                shaftPanel.Invalidate()
                If doorOpenPct >= 1.0F Then doorPhase = 2 : tmrDoor.Interval = 1000
            Case 2
                doorPhase = 3 : tmrDoor.Interval = 120
                AddLog("Doors closing at F" & currentFloor)
            Case 3
                doorOpenPct = Math.Max(0.0F, doorOpenPct - 0.2F)
                shaftPanel.Invalidate()
                If doorOpenPct <= 0 Then
                    tmrDoor.Stop() : tmrDoor.Interval = 600
                    doorPhase = 0 : doorsOpen = False : isMoving = False
                    RefreshStatus()
                    If Not isEmergency Then StartMoving()
                End If
        End Select
        UpdateDotIndicators()
    End Sub

    ' ── Emergency ──────────────────────────────────────────
    Private Sub btnEmergency_Click(sender As Object, e As EventArgs) Handles btnEmergency.Click
        If isEmergency Then
            ' Resume
            isEmergency = False
            btnEmergency.Text = "  ⚠  EMERGENCY STOP"
            btnEmergency.BackColor = Color.FromArgb(140, 20, 20)
            btnEmergency.ForeColor = Color.FromArgb(255, 80, 80)
            btnEmergency.FlatAppearance.BorderColor = Color.FromArgb(200, 50, 50)
            AddLog("▶ System resumed from emergency stop")
            RefreshStatus()
            If targetQueue.Count > 0 Then StartMoving()
        Else
            isEmergency = True
            tmrMove.Stop() : tmrDoor.Stop() : isMoving = False
            btnEmergency.Text = "  ▶  RESUME SYSTEM"
            btnEmergency.BackColor = Color.FromArgb(20, 80, 20)
            btnEmergency.ForeColor = Color.FromArgb(80, 220, 80)
            btnEmergency.FlatAppearance.BorderColor = Color.FromArgb(50, 180, 50)
            AddLog("⚠ EMERGENCY STOP ACTIVATED")
            shaftPanel.Invalidate()
            RefreshStatus()
        End If
    End Sub

    ' ── Hall Buttons ───────────────────────────────────────
    Private Sub HallUpButton_Click(sender As Object, e As EventArgs)
        If isEmergency Then AddLog("System in emergency stop!") : Return
        Dim btn As Button = DirectCast(sender, Button)
        Dim floor As Integer = CInt(btn.Tag)
        AddLog("Floor " & floor & " UP button pressed")
        btn.BackColor = Color.FromArgb(29, 158, 117)
        QueueFloor(floor)
    End Sub

    Private Sub HallDownButton_Click(sender As Object, e As EventArgs)
        If isEmergency Then AddLog("System in emergency stop!") : Return
        Dim btn As Button = DirectCast(sender, Button)
        Dim floor As Integer = CInt(btn.Tag)
        AddLog("Floor " & floor & " DOWN button pressed")
        btn.BackColor = Color.FromArgb(180, 60, 30)
        QueueFloor(floor)
    End Sub

    Private Sub CallButton_Click(sender As Object, e As EventArgs)
        If isEmergency Then AddLog("System in emergency stop!") : Return
        Dim btn As Button = DirectCast(sender, Button)
        Dim floor As Integer = CInt(btn.Tag)
        AddLog("Call button pressed: Floor " & floor)
        btn.BackColor = Color.FromArgb(30, 80, 150)
        QueueFloor(floor)
    End Sub

    Private Sub ClearFloorButtons(floor As Integer)
        If upButtons(floor) IsNot Nothing Then upButtons(floor).BackColor = Color.FromArgb(12, 36, 24)
        If downButtons(floor) IsNot Nothing Then downButtons(floor).BackColor = Color.FromArgb(36, 14, 10)
        If callButtons(floor) IsNot Nothing Then callButtons(floor).BackColor = Color.FromArgb(16, 28, 52)
    End Sub

    ' ── Dot Indicators ─────────────────────────────────────
    Private Sub UpdateDotIndicators()
        For f As Integer = 1 To TOTAL_FLOORS
            If dotPanels(f) Is Nothing Then Continue For
            dotPanels(f).Invalidate()
        Next
    End Sub

    Private Sub DotPanel_Paint(sender As Object, e As PaintEventArgs)
        Dim pnl As Panel = DirectCast(sender, Panel)
        Dim floor As Integer = CInt(pnl.Tag)
        Dim g As Graphics = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        g.Clear(Color.Transparent)

        ' Dot 1: lit if elevator is on this floor
        Dim dot1 As Color = If(currentFloor = floor, Color.FromArgb(80, 200, 160), Color.FromArgb(20, 50, 40))
        ' Dot 2: lit if doors open on this floor
        Dim dot2 As Color = If(currentFloor = floor AndAlso doorsOpen, Color.FromArgb(80, 160, 220), Color.FromArgb(20, 40, 60))

        g.FillEllipse(New SolidBrush(dot1), 0, 2, 8, 8)
        g.FillEllipse(New SolidBrush(dot2), 12, 2, 8, 8)
    End Sub

    ' ── Status ─────────────────────────────────────────────
    Private Sub RefreshStatus()
        ' Big floor display
        lblBigFloor.Text = currentFloor.ToString()

        lblFloorVal.Text = "Floor " & currentFloor

        If isEmergency Then
            lblDirectionVal.Text = "STOPPED"
            lblDirectionVal.ForeColor = Color.FromArgb(255, 80, 80)
            lblStatusVal.Text = "EMERGENCY"
            lblStatusVal.ForeColor = Color.FromArgb(255, 80, 80)
        ElseIf isMoving Then
            Dim pk As Integer = If(targetQueue.Count > 0, targetQueue.Peek(), currentFloor)
            lblDirectionVal.Text = If(pk > currentFloor, "▲  UP", "▼  DOWN")
            lblDirectionVal.ForeColor = Color.FromArgb(80, 180, 220)
            lblStatusVal.Text = "MOVING"
            lblStatusVal.ForeColor = Color.FromArgb(220, 180, 60)
        Else
            lblDirectionVal.Text = "IDLE"
            lblDirectionVal.ForeColor = Color.FromArgb(140, 160, 190)
            lblStatusVal.Text = "READY"
            lblStatusVal.ForeColor = Color.FromArgb(80, 200, 120)
        End If

        If passengerInElevator Then
            lblPassengerVal.Text = "In elevator → F" & passengerDest
            lblPassengerVal.ForeColor = Color.FromArgb(80, 200, 120)
        Else
            lblPassengerVal.Text = "Waiting at F" & passengerFloor
            lblPassengerVal.ForeColor = Color.FromArgb(220, 160, 60)
        End If

        UpdateDotIndicators()
        shaftPanel.Invalidate()
    End Sub

    Private Sub tmrClock_Tick(sender As Object, e As EventArgs) Handles tmrClock.Tick
        lblClock.Text = DateTime.Now.ToString("H:mm:ss tt")
        RefreshStatus()
    End Sub

    ' ── Log ────────────────────────────────────────────────
    Private Sub AddLog(msg As String)
        Dim ts As String = DateTime.Now.ToString("HH:mm:ss")
        logBox.Items.Insert(0, "[" & ts & "]  " & msg)
        If logBox.Items.Count > 80 Then logBox.Items.RemoveAt(logBox.Items.Count - 1)
    End Sub

    ' ── Shaft Drawing ──────────────────────────────────────
    Private Sub shaftPanel_Paint(sender As Object, e As PaintEventArgs) Handles shaftPanel.Paint
        Dim g As Graphics = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        Dim shaftH As Integer = TOTAL_FLOORS * FLOOR_HEIGHT
        g.Clear(Color.FromArgb(10, 14, 26))

        ' Floor bands
        For f As Integer = 1 To TOTAL_FLOORS
            Dim topY As Integer = shaftH - f * FLOOR_HEIGHT
            Dim bandClr As Color = If(f Mod 2 = 0,
                Color.FromArgb(14, 20, 38),
                Color.FromArgb(10, 16, 30))
            g.FillRectangle(New SolidBrush(bandClr), 0, topY, SHAFT_WIDTH, FLOOR_HEIGHT)
            g.DrawLine(New Pen(Color.FromArgb(25, 55, 100), 1), 0, topY, SHAFT_WIDTH, topY)

            ' Passenger waiting is drawn in a separate panel outside the shaft
        Next

        ' Rail lines
        Dim railX1 As Integer = SHAFT_WIDTH \ 2 - 4
        Dim railX2 As Integer = SHAFT_WIDTH \ 2 + 4
        g.DrawLine(New Pen(Color.FromArgb(25, 55, 100), 2), railX1, 0, railX1, shaftH)
        g.DrawLine(New Pen(Color.FromArgb(25, 55, 100), 2), railX2, 0, railX2, shaftH)

        ' Elevator car
        Dim carX As Integer = (SHAFT_WIDTH - CAR_WIDTH) \ 2
        Dim bodyClr As Color = If(isEmergency, Color.FromArgb(60, 10, 10), Color.FromArgb(14, 36, 62))
        Dim bdrClr As Color = If(isEmergency, Color.FromArgb(200, 50, 50), Color.FromArgb(50, 120, 200))

        ' Car shadow
        g.FillRectangle(New SolidBrush(Color.FromArgb(8, 0, 0, 0)), carX + 3, carY + 3, CAR_WIDTH, CAR_HEIGHT)
        ' Car body
        g.FillRectangle(New SolidBrush(bodyClr), carX, carY, CAR_WIDTH, CAR_HEIGHT)
        g.DrawRectangle(New Pen(bdrClr, 2), carX, carY, CAR_WIDTH, CAR_HEIGHT)

        ' Door panels
        Dim halfW As Integer = (CAR_WIDTH - 6) \ 2
        Dim doorW As Integer = CInt(halfW * (1 - doorOpenPct))
        Dim doorTop As Integer = carY + 6
        Dim doorH As Integer = CAR_HEIGHT - 8

        If doorW > 0 Then
            ' Left door
            Dim lgb As New Drawing2D.LinearGradientBrush(
                New Rectangle(carX + 2, doorTop, doorW, doorH),
                Color.FromArgb(30, 80, 140), Color.FromArgb(20, 55, 100),
                Drawing2D.LinearGradientMode.Horizontal)
            g.FillRectangle(lgb, carX + 2, doorTop, doorW, doorH)
            g.DrawRectangle(New Pen(Color.FromArgb(60, 130, 200), 1), carX + 2, doorTop, doorW, doorH)

            ' Right door
            Dim rgb2 As New Drawing2D.LinearGradientBrush(
                New Rectangle(carX + CAR_WIDTH - 2 - doorW, doorTop, doorW, doorH),
                Color.FromArgb(20, 55, 100), Color.FromArgb(30, 80, 140),
                Drawing2D.LinearGradientMode.Horizontal)
            g.FillRectangle(rgb2, carX + CAR_WIDTH - 2 - doorW, doorTop, doorW, doorH)
            g.DrawRectangle(New Pen(Color.FromArgb(60, 130, 200), 1), carX + CAR_WIDTH - 2 - doorW, doorTop, doorW, doorH)
        End If

        ' Passenger inside (draw person icon)
        If passengerInElevator Then
            Dim personHIn As Integer = CInt(CAR_HEIGHT * 0.45)
            Dim centerXIn As Integer = carX + CAR_WIDTH \ 2
            Dim bottomYIn As Integer = carY + CAR_HEIGHT - 8
            DrawPerson(g, centerXIn, bottomYIn, personHIn, Color.FromArgb(80, 220, 140))
        End If
    End Sub

    ' Draw a simple person icon centered at (centerX, bottomY) with total height
    Private Sub DrawPerson(g As Graphics, centerX As Integer, bottomY As Integer, height As Integer, clr As Color)
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        ' Colors: skin tone and clothing
        Dim skin As Color = Color.FromArgb(250, 210, 170)
        Dim cloth As Color = clr

        Dim headR As Integer = Math.Max(3, CInt(height * 0.2))
        Dim headCX As Integer = centerX
        Dim headCY As Integer = bottomY - height + headR

        ' torso rectangle
        Dim torsoTopY As Integer = headCY + headR + 2
        Dim torsoBottomY As Integer = bottomY - CInt(height * 0.3)
        Dim torsoW As Integer = CInt(height * 0.38)

        ' legs box height
        Dim legTopY As Integer = torsoBottomY
        Dim legBottomY As Integer = bottomY

        Using brSkin As New SolidBrush(skin), brCloth As New SolidBrush(cloth), pen As New Pen(Color.FromArgb(30, 30, 30), 1)
            ' head
            g.FillEllipse(brSkin, headCX - headR, headCY - headR, headR * 2, headR * 2)
            g.DrawEllipse(pen, headCX - headR, headCY - headR, headR * 2, headR * 2)

            ' torso (shirt)
            Dim torsoRect As New Rectangle(centerX - torsoW \ 2, torsoTopY, torsoW, Math.Max(4, torsoBottomY - torsoTopY))
            g.FillRectangle(brCloth, torsoRect)
            g.DrawRectangle(pen, torsoRect)

            ' arms (filled rounded rectangles)
            Dim armW = Math.Max(4, CInt(torsoW * 0.28))
            Dim armH = Math.Max(4, CInt((torsoRect.Height) * 0.28))
            g.FillRectangle(brCloth, centerX - torsoW \ 2 - armW, torsoTopY + 4, armW, armH)
            g.FillRectangle(brCloth, centerX + torsoW \ 2, torsoTopY + 4, armW, armH)

            ' neck (small skin rect)
            g.FillRectangle(brSkin, centerX - 2, torsoTopY - 2, 4, 4)

            ' legs (pants)
            Dim legW As Integer = Math.Max(6, CInt(torsoW * 0.46))
            Dim gap As Integer = 6
            Dim leftLegRect As New Rectangle(centerX - gap - legW, legTopY, legW, legBottomY - legTopY)
            Dim rightLegRect As New Rectangle(centerX + gap, legTopY, legW, legBottomY - legTopY)
            Using brLeg As New SolidBrush(Color.FromArgb(60, 60, 80))
                g.FillRectangle(brLeg, leftLegRect)
                g.FillRectangle(brLeg, rightLegRect)
                g.DrawRectangle(pen, leftLegRect)
                g.DrawRectangle(pen, rightLegRect)
            End Using
        End Using
    End Sub

    ' ── Build UI ───────────────────────────────────────────
    Private Sub BuildUI()
        Dim shaftH As Integer = TOTAL_FLOORS * FLOOR_HEIGHT   ' 360
        Dim formW As Integer = Me.ClientSize.Width            ' ~694
        Dim formH As Integer = Me.ClientSize.Height           ' ~481

        ' ── LEFT PANEL (shaft area) ────────────────────────
        Dim leftPnl As New Panel()
        leftPnl.Location = New Point(0, 0)
        leftPnl.Size = New Size(360, formH)
        leftPnl.BackColor = Color.FromArgb(14, 20, 38)
        Me.Controls.Add(leftPnl)

        ' Title strip on left panel
        Dim lblTitle As New Label()
        lblTitle.Text = "ElevatorSim v1.0  —  4 Floor Elevator Control System"
        lblTitle.Font = New Font("Courier New", 8, FontStyle.Bold)
        lblTitle.ForeColor = Color.FromArgb(80, 120, 200)
        lblTitle.Location = New Point(0, 8)
        lblTitle.Size = New Size(260, 18)
        lblTitle.TextAlign = ContentAlignment.MiddleCenter
        leftPnl.Controls.Add(lblTitle)

        ' Shaft panel centred in left area
        Dim shaftLeft As Integer = 100   ' x offset inside leftPnl
        Dim shaftTop As Integer = 38
        shaftPanel = New Panel()
        shaftPanel.Location = New Point(shaftLeft, shaftTop)
        shaftPanel.Size = New Size(SHAFT_WIDTH, shaftH)
        shaftPanel.BackColor = Color.FromArgb(10, 14, 26)
        shaftPanel.BorderStyle = BorderStyle.None
        leftPnl.Controls.Add(shaftPanel)
        ' Enable double buffering to reduce flicker
        EnableDoubleBuffer(shaftPanel)

        ' Thin border around shaft
        Dim shaftBorder As New Panel()
        shaftBorder.Location = New Point(shaftLeft - 1, shaftTop - 1)
        shaftBorder.Size = New Size(SHAFT_WIDTH + 2, shaftH + 2)
        shaftBorder.BackColor = Color.FromArgb(30, 60, 110)
        shaftBorder.SendToBack()
        leftPnl.Controls.Add(shaftBorder)
        shaftPanel.BringToFront()

        ' ── Hall buttons + floor labels (left of shaft) ────
        For f As Integer = 1 To TOTAL_FLOORS
            Dim midY As Integer = shaftTop + shaftH - (f - 1) * FLOOR_HEIGHT - FLOOR_HEIGHT \ 2

            ' Floor number label (large, left margin)
            Dim lblNum As New Label()
            lblNum.Text = f.ToString()
            lblNum.Font = New Font("Courier New", 14, FontStyle.Bold)
            lblNum.ForeColor = Color.FromArgb(160, 190, 230)
            lblNum.Location = New Point(8, midY - 14)
            lblNum.Size = New Size(28, 28)
            lblNum.TextAlign = ContentAlignment.MiddleCenter
            leftPnl.Controls.Add(lblNum)

            ' Floor name label (small, below number)
            Dim lblName As New Label()
            lblName.Text = FloorNames(f - 1)
            lblName.Font = New Font("Courier New", 6, FontStyle.Bold)
            lblName.ForeColor = Color.FromArgb(70, 100, 150)
            lblName.Location = New Point(6, midY + 12)
            lblName.Size = New Size(32, 14)
            lblName.TextAlign = ContentAlignment.MiddleCenter
            leftPnl.Controls.Add(lblName)

            ' UP button (all floors except top)
            If f < TOTAL_FLOORS Then
                Dim ub As New Button()
                ub.Text = "▲" : ub.Tag = f
                ub.Size = New Size(36, 22)
                ub.Location = New Point(shaftLeft - 50, midY - 26)
                ub.BackColor = Color.FromArgb(12, 36, 24)
                ub.ForeColor = Color.FromArgb(80, 200, 140)
                ub.FlatStyle = FlatStyle.Flat
                ub.FlatAppearance.BorderColor = Color.FromArgb(30, 130, 80)
                ub.Font = New Font("Courier New", 8, FontStyle.Bold)
                ub.Cursor = Cursors.Hand
                AddHandler ub.Click, AddressOf HallUpButton_Click
                upButtons(f) = ub
                leftPnl.Controls.Add(ub)
            End If

            ' DOWN button (all floors except ground)
            If f > 1 Then
                Dim db As New Button()
                db.Text = "▼" : db.Tag = f
                db.Size = New Size(36, 22)
                db.Location = New Point(shaftLeft - 50, midY + 4)
                db.BackColor = Color.FromArgb(36, 14, 10)
                db.ForeColor = Color.FromArgb(220, 100, 70)
                db.FlatStyle = FlatStyle.Flat
                db.FlatAppearance.BorderColor = Color.FromArgb(160, 60, 30)
                db.Font = New Font("Courier New", 8, FontStyle.Bold)
                db.Cursor = Cursors.Hand
                AddHandler db.Click, AddressOf HallDownButton_Click
                downButtons(f) = db
                leftPnl.Controls.Add(db)
            End If

            ' Floor label (right of shaft)
            Dim lblFR As New Label()
            lblFR.Text = f.ToString() & " LEV"
            If f = 1 Then lblFR.Text = "1 GRD"
            If f = TOTAL_FLOORS Then lblFR.Text = f.ToString() & " TOP"
            lblFR.Font = New Font("Courier New", 7, FontStyle.Bold)
            lblFR.ForeColor = Color.FromArgb(60, 90, 140)
            lblFR.Location = New Point(shaftLeft + SHAFT_WIDTH + 4, midY - 8)
            lblFR.Size = New Size(50, 16)
            leftPnl.Controls.Add(lblFR)

            ' Dot indicators (right of shaft)
            Dim dp As New Panel()
            dp.Tag = f
            dp.Size = New Size(24, 14)
            dp.Location = New Point(shaftLeft + SHAFT_WIDTH + 4, midY + 8)
            dp.BackColor = Color.Transparent
            AddHandler dp.Paint, AddressOf DotPanel_Paint
            dotPanels(f) = dp
            leftPnl.Controls.Add(dp)

            ' Person panel (outside the shaft, right side)
            Dim pp As New Panel()
            pp.Tag = f
            Dim personW As Integer = 40
            Dim personH As Integer = Math.Min(60, FLOOR_HEIGHT \ 3)
            pp.Size = New Size(personW, personH + 6)
            ' place to the right of the shaft (after dot indicators / labels)
            ' move person panel further right for more space
            pp.Location = New Point(shaftLeft + SHAFT_WIDTH + 72, midY - personH)
            pp.BackColor = Color.Transparent
            pp.Visible = False ' don't show until spawned
            AddHandler pp.Paint, AddressOf PersonPanel_Paint
            personPanels(f) = pp
            leftPnl.Controls.Add(pp)
            EnableDoubleBuffer(pp)
        Next

        ' ── RIGHT PANEL ────────────────────────────────────
        Dim rightPnl As New Panel()
        rightPnl.Location = New Point(362, 0)
        rightPnl.Size = New Size(formW - 362, formH)
        rightPnl.BackColor = Color.FromArgb(18, 24, 42)
        Me.Controls.Add(rightPnl)

        Dim rx As Integer = 16   ' right panel inner x margin
        Dim ry As Integer = 10

        ' Clock top-right
        lblClock = New Label()
        lblClock.Text = DateTime.Now.ToString("H:mm:ss tt")
        lblClock.Font = New Font("Courier New", 9)
        lblClock.ForeColor = Color.FromArgb(100, 130, 180)
        lblClock.Size = New Size(140, 18)
        lblClock.Location = New Point(rightPnl.Width - 150, ry)
        lblClock.TextAlign = ContentAlignment.MiddleRight
        rightPnl.Controls.Add(lblClock)

        ' ── ELEVATOR DISPLAY section ───────────────────────
        Dim lblDispHead As New Label()
        lblDispHead.Text = "ELEVATOR DISPLAY"
        lblDispHead.Font = New Font("Courier New", 9, FontStyle.Bold)
        lblDispHead.ForeColor = Color.FromArgb(200, 60, 60)
        lblDispHead.Location = New Point(rx, ry)
        lblDispHead.Size = New Size(200, 18)
        rightPnl.Controls.Add(lblDispHead)
        ry += 24

        ' Big floor number box
        Dim bigBox As New Panel()
        bigBox.Location = New Point(rx, ry)
        bigBox.Size = New Size(rightPnl.Width - rx - 16, 52)
        bigBox.BackColor = Color.Black
        rightPnl.Controls.Add(bigBox)

        lblBigFloor = New Label()
        lblBigFloor.Text = currentFloor.ToString()
        lblBigFloor.Font = New Font("Courier New", 28, FontStyle.Bold)
        lblBigFloor.ForeColor = Color.FromArgb(60, 180, 220)
        lblBigFloor.Dock = DockStyle.Fill
        lblBigFloor.TextAlign = ContentAlignment.MiddleCenter
        bigBox.Controls.Add(lblBigFloor)
        ry += 60

        ' Status rows
        Dim statusW As Integer = rightPnl.Width - rx - 16
        rightPnl.Controls.Add(MkStatusRow("Floor:", lblFloorVal, New Point(rx, ry), statusW, Color.FromArgb(80, 200, 120))) : ry += 22
        rightPnl.Controls.Add(MkStatusRow("Direction:", lblDirectionVal, New Point(rx, ry), statusW, Color.FromArgb(140, 160, 190))) : ry += 22
        rightPnl.Controls.Add(MkStatusRow("Status:", lblStatusVal, New Point(rx, ry), statusW, Color.FromArgb(80, 200, 120))) : ry += 22
        rightPnl.Controls.Add(MkStatusRow("Passenger:", lblPassengerVal, New Point(rx, ry), statusW, Color.FromArgb(220, 160, 60))) : ry += 28

        ' Separator
        rightPnl.Controls.Add(MkSep(ry, rightPnl.Width - rx - 16, rx))
        ry += 10

        ' ── CALL ELEVATOR section ──────────────────────────
        Dim lblCallHead As New Label()
        lblCallHead.Text = "CALL ELEVATOR TO FLOOR"
        lblCallHead.Font = New Font("Courier New", 9, FontStyle.Bold)
        lblCallHead.ForeColor = Color.FromArgb(200, 60, 60)
        lblCallHead.Location = New Point(rx, ry)
        lblCallHead.Size = New Size(240, 18)
        rightPnl.Controls.Add(lblCallHead)
        ry += 24

        ' 2x2 grid of call buttons
        Dim btnW As Integer = (rightPnl.Width - rx - 16 - 8) \ 2
        Dim btnH As Integer = 32
        For f As Integer = 1 To TOTAL_FLOORS
            Dim col As Integer = (f - 1) Mod 2
            Dim row As Integer = (f - 1) \ 2
            Dim cb As New Button()
            cb.Text = "Floor " & f : cb.Tag = f
            cb.Size = New Size(btnW, btnH)
            cb.Location = New Point(rx + col * (btnW + 8), ry + row * (btnH + 6))
            cb.BackColor = Color.FromArgb(16, 28, 52)
            cb.ForeColor = Color.FromArgb(160, 200, 240)
            cb.FlatStyle = FlatStyle.Flat
            cb.FlatAppearance.BorderColor = Color.FromArgb(40, 80, 160)
            cb.Font = New Font("Courier New", 9, FontStyle.Bold)
            cb.Cursor = Cursors.Hand
            AddHandler cb.Click, AddressOf CallButton_Click
            callButtons(f) = cb
            rightPnl.Controls.Add(cb)
        Next
        ry += 2 * (btnH + 6) + 6

        ' ── MANUAL PASSENGER CONTROLS ─────────────────────
        Dim lblManualHead As New Label()
        lblManualHead.Text = "MANUAL PASSENGER"
        lblManualHead.Font = New Font("Courier New", 9, FontStyle.Bold)
        lblManualHead.ForeColor = Color.FromArgb(200, 60, 60)
        lblManualHead.Location = New Point(rx, ry)
        lblManualHead.Size = New Size(240, 18)
        rightPnl.Controls.Add(lblManualHead)
        ry += 22

        cmbSpawnFloor = New ComboBox()
        cmbSpawnFloor.DropDownStyle = ComboBoxStyle.DropDownList
        cmbSpawnFloor.Font = New Font("Courier New", 9)
        For i As Integer = 1 To TOTAL_FLOORS
            cmbSpawnFloor.Items.Add(i)
        Next
        cmbSpawnFloor.SelectedIndex = 0
        cmbSpawnFloor.Size = New Size(80, 24)
        cmbSpawnFloor.Location = New Point(rx, ry)
        rightPnl.Controls.Add(cmbSpawnFloor)

        cmbPassengerDest = New ComboBox()
        cmbPassengerDest.DropDownStyle = ComboBoxStyle.DropDownList
        cmbPassengerDest.Font = New Font("Courier New", 9)
        For i As Integer = 1 To TOTAL_FLOORS
            cmbPassengerDest.Items.Add(i)
        Next
        cmbPassengerDest.SelectedIndex = Math.Min(1, cmbPassengerDest.Items.Count - 1)
        cmbPassengerDest.Size = New Size(80, 24)
        cmbPassengerDest.Location = New Point(rx + 92, ry)
        rightPnl.Controls.Add(cmbPassengerDest)

        btnSpawnPassenger = New Button()
        btnSpawnPassenger.Text = "Spawn Passenger"
        btnSpawnPassenger.Size = New Size(160, 28)
        btnSpawnPassenger.Location = New Point(rx + 184, ry - 2)
        btnSpawnPassenger.BackColor = Color.FromArgb(30, 60, 40)
        btnSpawnPassenger.ForeColor = Color.FromArgb(180, 220, 180)
        btnSpawnPassenger.FlatStyle = FlatStyle.Flat
        btnSpawnPassenger.Cursor = Cursors.Hand
        AddHandler btnSpawnPassenger.Click, AddressOf btnSpawnPassenger_Click
        rightPnl.Controls.Add(btnSpawnPassenger)
        ry += 34

        btnBoardPassenger = New Button()
        btnBoardPassenger.Text = "Board / Exit Passenger"
        btnBoardPassenger.Size = New Size(rightPnl.Width - rx - 16, 32)
        btnBoardPassenger.Location = New Point(rx, ry)
        btnBoardPassenger.BackColor = Color.FromArgb(40, 40, 80)
        btnBoardPassenger.ForeColor = Color.FromArgb(200, 200, 240)
        btnBoardPassenger.FlatStyle = FlatStyle.Flat
        btnBoardPassenger.Cursor = Cursors.Hand
        AddHandler btnBoardPassenger.Click, AddressOf btnBoardPassenger_Click
        rightPnl.Controls.Add(btnBoardPassenger)
        ry += 38

        ' ── EMERGENCY button ───────────────────────────────
        btnEmergency = New Button()
        btnEmergency.Text = "  ⚠  EMERGENCY STOP"
        btnEmergency.Size = New Size(rightPnl.Width - rx - 16, 36)
        btnEmergency.Location = New Point(rx, ry)
        btnEmergency.BackColor = Color.FromArgb(140, 20, 20)
        btnEmergency.ForeColor = Color.FromArgb(255, 80, 80)
        btnEmergency.FlatStyle = FlatStyle.Flat
        btnEmergency.FlatAppearance.BorderColor = Color.FromArgb(200, 50, 50)
        btnEmergency.Font = New Font("Courier New", 10, FontStyle.Bold)
        btnEmergency.Cursor = Cursors.Hand
        btnEmergency.TextAlign = ContentAlignment.MiddleLeft
        rightPnl.Controls.Add(btnEmergency)
        ry += 44

        ' Separator
        rightPnl.Controls.Add(MkSep(ry, rightPnl.Width - rx - 16, rx))
        ry += 10

        ' ── SYSTEM LOG ─────────────────────────────────────
        Dim lblLogHead As New Label()
        lblLogHead.Text = "SYSTEM LOG"
        lblLogHead.Font = New Font("Courier New", 9, FontStyle.Bold)
        lblLogHead.ForeColor = Color.FromArgb(200, 60, 60)
        lblLogHead.Location = New Point(rx, ry)
        lblLogHead.Size = New Size(150, 18)
        rightPnl.Controls.Add(lblLogHead)
        ry += 22

        logBox = New ListBox()
        logBox.Location = New Point(rx, ry)
        logBox.Size = New Size(rightPnl.Width - rx - 16, formH - ry - 10)
        logBox.BackColor = Color.FromArgb(10, 14, 26)
        logBox.ForeColor = Color.FromArgb(110, 160, 210)
        logBox.Font = New Font("Courier New", 8)
        logBox.BorderStyle = BorderStyle.FixedSingle
        logBox.DrawMode = DrawMode.OwnerDrawFixed
        logBox.ItemHeight = 16
        AddHandler logBox.DrawItem, AddressOf LogBox_DrawItem
        rightPnl.Controls.Add(logBox)
    End Sub

    ' ── Status row helper ──────────────────────────────────
    Private Function MkStatusRow(labelTxt As String, ByRef valLbl As Label,
                                  loc As Point, totalW As Integer,
                                  valColor As Color) As Panel
        Dim row As New Panel()
        row.Location = loc
        row.Size = New Size(totalW, 20)
        row.BackColor = Color.Transparent

        Dim lbl As New Label()
        lbl.Text = labelTxt
        lbl.Font = New Font("Courier New", 9)
        lbl.ForeColor = Color.FromArgb(80, 100, 140)
        lbl.Location = New Point(0, 0)
        lbl.Size = New Size(90, 20)
        row.Controls.Add(lbl)

        valLbl = New Label()
        valLbl.Text = ""
        valLbl.Font = New Font("Courier New", 9, FontStyle.Bold)
        valLbl.ForeColor = valColor
        valLbl.Location = New Point(92, 0)
        valLbl.Size = New Size(totalW - 92, 20)
        valLbl.TextAlign = ContentAlignment.MiddleRight
        row.Controls.Add(valLbl)

        Return row
    End Function

    Private Function MkSep(y As Integer, w As Integer, x As Integer) As Label
        Return New Label() With {
            .Size = New Size(w, 1),
            .Location = New Point(x, y),
            .BackColor = Color.FromArgb(30, 55, 100)
        }
    End Function

    ' ── Custom log draw ────────────────────────────────────
    Private Sub LogBox_DrawItem(sender As Object, e As DrawItemEventArgs)
        If e.Index < 0 OrElse e.Index >= logBox.Items.Count Then Return
        e.DrawBackground()
        Dim msg As String = logBox.Items(e.Index).ToString()
        Dim fc As Color = Color.FromArgb(110, 160, 210)
        If msg.Contains("EMERGENCY") OrElse msg.Contains("⚠") Then
            fc = Color.FromArgb(255, 80, 80)
        ElseIf msg.Contains("Arrived") OrElse msg.Contains("boarded") OrElse
               msg.Contains("exited") OrElse msg.Contains("resumed") OrElse
               msg.Contains("initialised") OrElse msg.Contains("✦") Then
            fc = Color.FromArgb(80, 200, 120)
        ElseIf msg.Contains("Moving") OrElse msg.Contains("pressed") OrElse
               msg.Contains("Doors") OrElse msg.Contains("Waiting") Then
            fc = Color.FromArgb(220, 180, 60)
        End If
        e.Graphics.DrawString(msg, e.Font, New SolidBrush(fc), e.Bounds)
        e.DrawFocusRectangle()
    End Sub

    Private Sub EnableDoubleBuffer(ctrl As Control)
        ' Reduce flicker by enabling double buffering via reflection (works for Panel)
        Try
            Dim prop = ctrl.GetType().GetProperty("DoubleBuffered", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.Public)
            If prop IsNot Nothing Then prop.SetValue(ctrl, True, Nothing)
        Catch
        End Try
    End Sub

End Class
