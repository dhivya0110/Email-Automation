# ----STEP 1 : LOAD THE EXCEL FROM ONE DRIVE----
$excelFilePath = "C:\Users\DHIVYAB\OneDrive - Capgemini\Daily Health checks\Tempalte - Operational Status Dashboard_V 0.2.xlsx"

# ----STEP 2: TEST FILE EXIST----
if (!(Test-Path $excelFilePath)) {
    Write-Host "Excel file not found: $excelFilePath"
    exit
}

# ----STEP 3: CHECK IF FILE WAS MODIFIED IN THE LAST 10 MINS----
$lastWriteTime = (Get-Item $excelFilePath).LastWriteTime
if ((Get-Date) - $lastWriteTime -gt (New-TimeSpan -Minutes 10)) {
    Write-Host "File not updated in the last 10 Mins. Last modified at: $lastWriteTime"
    exit
}

# ----STEP 4: OPEN EXCEL IN HIDDEN MODE----
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelFilePath)
#$sheet = $workbook.Sheets.Item(1)

$sheetCount = $workbook.Sheets.Count
$sheet = $workbook.Sheets.Item($sheetCount)  #working in this area!


# ----STEP 5: DEFINE CELL RANGES TO COPY----
$rangeProd = $sheet.Range("A1:L22") #need to work
$rangeQA = $sheet.Range("M1:X22")

# ---- STEP 6: START OUTLOOK AND CREATE A DRAFT EMAIL---- 
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.Display()

# ---- STEP 7: ALLOW TIME FOR OUTLOOK TO LOAD
Start-Sleep -Seconds 2
$inspector = $mail.GetInspector()
$editor = $inspector.WordEditor.Application.Selection

# ----STEP 8: CURSOR TO THE TOP AND BEGIN DRAFTING MAIL---- 
$editor.HomeKey(6) 
$editor.TypeText("Hello All,")
$editor.TypeParagraph()
$editor.TypeText("Greetings for the day.")
$editor.TypeParagraph()
$editor.TypeParagraph()
$editor.TypeText("Please find below the health check report for Prod and QA.")
$editor.TypeParagraph()
$editor.TypeParagraph()

# ----STEP 9: ADD PROD HEADING WITH BOLD AND UNDERLINE----
$editor.Font.Bold = $true
$editor.Font.Underline = 1 
$editor.TypeText("PROD:")
$editor.Font.Bold = $false
$editor.Font.Underline = 0
$editor.TypeParagraph()

# ----STEP 10: COPY PROD RANGE AS IMAGE AND PASTE----
$rangeProd.CopyPicture([Microsoft.Office.Interop.Excel.XlPictureAppearance]::xlScreen,
                       [Microsoft.Office.Interop.Excel.XlCopyPictureFormat]::xlBitmap)
Start-Sleep -Milliseconds 700
$editor.Paste()
$editor.TypeParagraph()
$editor.TypeParagraph()

# ----STEP 11: ADD QA HEANDING WITH BOLD AND UNDERLINE----
$editor.Font.Bold = $true
$editor.Font.Underline = 1
$editor.TypeText("QA:")
$editor.Font.Bold = $false
$editor.Font.Underline = 0
$editor.TypeParagraph()

# ----STEP 12: COPY QA RANGE AS IMAGE AND PASTE----
$rangeQA.CopyPicture([Microsoft.Office.Interop.Excel.XlPictureAppearance]::xlScreen,
                     [Microsoft.Office.Interop.Excel.XlCopyPictureFormat]::xlBitmap)
Start-Sleep -Milliseconds 700
$editor.Paste()
$editor.TypeParagraph()

# ----STEP 13: SET EMAIL META INFO----
$mail.Subject = "Health Check Report - $(Get-Date -Format 'dd-MM-yyyy')"
$mail.To = "dhivyalakshmi.b@capgemini.com"
$mail.CC = "dhivyalakshmi.b@capgemini.com"
$mail.SentOnBehalfOfName = "dhivyalakshmi.b@capgemini.com" 

# ----STEP 14: CLEAN UP THE EXCEL----
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null




