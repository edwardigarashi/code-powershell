#SCreate and get my Excel Obj
$x1 = New-Object -comobject Excel.Application
$UserWorkBook = $x1.Workbooks.Open("C:\Users\edwar\Desktop\class schedule.xlsx")

#Select first Sheet
$UserWorksheet = $UserWorkBook.Worksheets.Item(1)
$UserWorksheet.activate()

#Copy the part of the sheet I want in the Email

$rgeSource=$UserWorksheet.range("D12","P39").Copy()

#create outlook Object
$Outlook = New-Object -comObject  Outlook.Application 
$Mail = $Outlook.CreateItem(0) 
$Mail.Recipients.Add("edward@igarashi.one") 

#Add the text part I want to display first
$Mail.Body = "My Comment on the Excel Spreadsheet" 

#Add Subject
$Mail.Subject = "Hellow"

#Then Copy the Excel using parameters to format it
$Mail.Getinspector.WordEditor.Range().PasteExcelTable($true,$false,$True)
#Then it becomes possible to insert text before
$wdDoc = $Mail.Getinspector.WordEditor
$wdRange = $wdDoc.Range()
$wdRange.InsertBefore("!")
$Mail.Display()

$Mail.Send()