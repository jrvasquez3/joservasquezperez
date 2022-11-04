<div>
    <ul class="nav">
        <li class="nav"><a href="google.com">Home</a></li>
        <li class="nav"><a href="google.com">About</a></li>
        <li class="nav"><a href="google.com">Contact</a></li>
    </ul>
</div>

<link rel="stylesheet" href="styles.css">

#
# KPI 
### Author: Jose R Vasquez Perez

# Table of Contents
1. [Enable Developer Tools](#enable-developer-tools)
2. [START SHIFT](#start-shift)
3. [END SHIFT - not done](#third-example)
4. [CURRENT DATA - not done](#fourth-examplehttpwwwfourthexamplecom)

#

I have made and created KPI for my past employer and thus, I want to share some of my knowledge to any reader curious about how to create KPIs using VBA. Below are some basic instructions I have created.

1. <b>Enable Developer Tools:</b>
# Enable Developer Tools

#
First Step before developing anything with *Visual Basic for Applications* would be by enabling developer mode inside the application where you want to develop the dashboard (Microsoft Excel for our purposes).

2. <b>Create Start Shift Button:</b>

# Start Shift
Add a *Command Button* into the sheet. Then, in VBA editor, if you want to allow the user to START SHIFT, do the following:

```vb
Sub start_shift()
'Author: Jose R Vasquez Perez

' Initializing variables
Dim last As Integer
Dim result As Range
Set result = Sheets("Shift Record").Columns(1).Find("Start Shift")
```

Afterwards, we would first like to ask the operator (or user) to enter the Shift number that they are working. Below is code I developed to do just that. Keep in mind that this is just one method to accomplish this and there are many other ways to do just this. The code below has the chance of "injection", however, for our purposes, this is okay.

```vb

'ask operator to enter current shift - JV
Range("X13").Value = InputBox("Please insert your shift number: 1 = first ; 2 = second ; 3 = third", "shift definition", "")
```

We will now add the current date to the shift record sheet. 

```vb

'Finds last row on Shift record sheet. Types current day of the week into the column named "comment" 
'Displays current day (M-Sunday) into Main Panel Sheet Tab - JV

last = Sheets("Shift Record").Cells(Rows.Count, 1).End(xlUp).Row + 4
day_of_week = Weekday(Date)
If day_of_week = 1 Then
    Sheets("Shift Record").Cells(last, 6).Value = "Sunday " & Range("X13").Value
    Sheets("panel").Range("W14").Value = "Sunday"
ElseIf day_of_week = 2 Then
    Sheets("Shift Record").Cells(last, 6).Value = "Monday " & Range("X13").Value
    Sheets("panel").Range("W14").Value = "Monday"
ElseIf day_of_week = 3 Then
    Sheets("Shift Record").Cells(last, 6).Value = "Tuesday " & Range("X13").Value
    Sheets("panel").Range("W14").Value = "Tuesday"
ElseIf day_of_week = 4 Then
    Sheets("Shift Record").Cells(last, 6).Value = "Wednesday " & Range("X13").Value
    Sheets("panel").Range("W14").Value = "Wednesday"
ElseIf day_of_week = 5 Then
    Sheets("Shift Record").Cells(last, 6).Value = "Thursday " & Range("X13").Value
    Sheets("panel").Range("W14").Value = "Thursday"
ElseIf day_of_week = 6 Then
    Sheets("Shift Record").Cells(last, 6).Value = "Friday " & Range("X13").Value
    Sheets("panel").Range("W14").Value = "Friday"
ElseIf day_of_week = 7 Then
    Sheets("Shift Record").Cells(last, 6).Value = "Saturday " & Range("X13").Value
    Sheets("panel").Range("W14").Value = "Saturday"

End If

```

As you can tell, this code uses repeated "Elseif" statements. There are obviously better methods to implement this, however I decided to use only "ElseIf" statements for readability.

The reason as to why I decided to do this is because I trained an operator to be able to "read" VBA code in the event that when I am gone, he would be able to "read" the code and understand what is going on. I opted for better readability for non-technical individuals. 

```vb

'Types "Start Shift", today's date and time into the last row - JV

last = Sheets("Shift Record").Cells(Rows.Count, 1).End(xlUp).Row + 4
Sheets("Shift Record").Cells(last, 1).Value = "Start Shift"

Sheets("Shift Record").Cells(last, 2).Value = Date

Sheets("Shift Record").Cells(last, 3).Value = Time

End Sub
```



