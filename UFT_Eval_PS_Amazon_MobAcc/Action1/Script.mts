systemutil.Run "chrome.exe","https://www.amazon.in/"
wait 2

Set orAmazon = Browser("Amazon").Page("AmazonPG")
varResult  = "xpath:=//div[@class='a-section a-spacing-none']/div/div/div/span[2]"

orAmazon.Link("Mobiles").Highlight
orAmazon.Link("Mobiles").Click
wait 3

orAmazon.Link("Mobiles & Accessories").Highlight
orAmazon.Link("Mobiles & Accessories").Click
wait 3


If Trim("Mobiles & Accessories") =  Trim(orAmazon.WebElement(varResult).GetROProperty("innertext"))Then
	orAmazon.WebElement(varResult).Highlight
	Print("Mobile Accessories Result: " & orAmazon.WebElement(varResult).GetROProperty("innertext"))
Else
	Print("Mobile Accessories Result is not displayed ")
	
End If


' Describe all product items in the amazon under the mobile and accessories
Set objDesc = Description.Create()
objDesc("micclass").Value = "WebElement"
'objDesc("class").Value = "a-size-base-plus a-spacing-none a-color-base a-text-normal" 
'objDesc("html tag").Value = "SPAN"
objDesc("xpath").Value = "//h2[@class='a-size-base-plus a-spacing-none a-color-base a-text-normal']" ' "//DIV[9]/SPAN[1]/DIV[1]/DIV[1]/DIV[2]/DIV[1]/A[1]/H2[1]/SPAN[1]"  '"//DIV[9]/SPAN[1]/DIV[1]/DIV[1]/DIV[2]/DIV[1]/A[1]/H2[1]/SPAN[1]"

'//div[@class='a-link-normal s-line-clamp-4 s-link-style a-text-normal']//h2[@class='a-size-base-plus a-spacing-none a-color-base a-text-normal']

' Get all product elements from the page
Set objProducts = Browser("CreationTime:=0").Page("index:=0").ChildObjects(objDesc)

' Count and print all product names
productCount = objProducts.Count
Print "Total number of products found: " & productCount

' Loop through all product names and print them
For i = 0 To productCount - 1
    productName = objProducts(i).GetROProperty("innertext")
    Print "Product " & (i+1) & ": " & productName
    
    Dim pName
    pName = objProducts(i).GetROProperty("innertext")
    Call fnWriteExcelResult(i)
    
Next

' Cleanup
Set objProducts = Nothing
Set objDesc = Nothing


Public Function fnWriteExcelResult(i)

' Create Excel Application Object
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False  ' To see Excel open

' Open the Excel Workbook
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\MBaladhandapani\Desktop\UFT_Eval_PS\MobileAccessoriesSearchResults.xlsx")

' Select the Sheet
Set objSheet = objWorkbook.Sheets("Sheet1")

Dim vRow
vRow = i+1

'objSheet.UsedRange.Offset(vRow, 1).ClearContents
'objSheet.UsedRange.Offset(vRow, 2).ClearContents

' Assuming row 1 contains headers
objSheet.Cells(vRow+1, 1).Value =  vRow ' write column SlNo
objSheet.Cells(vRow+1, 2).Value =  pName ' write column MobileandAccessories
Print "vRow " & vRow & ":"  &"Product"  & pName

'Save
objWorkbook.Save
objWorkbook.Close False
objExcel.Quit

Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing

	
End Function	

Call fnReadExcelResult()

Public Function fnReadExcelResult()

' Create Excel Application Object
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False  'To open the excel visually

' Open  Excel Workbook
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\MBaladhandapani\Desktop\UFT_Eval_PS\MobileAccessoriesSearchResults.xlsx")

' Obj sheet
Set objSheet = objWorkbook.Sheets("Sheet1")

' Get the last used row 
lastRow = objSheet.UsedRange.Rows.Count

' Loop through each row and read data

For i = 2 To lastRow  ' row 1 contains headers
    vSlNo = objSheet.Cells(i, 1).Value  ' Read column SlNo
    vMobileandAccessories = objSheet.Cells(i, 2).Value  ' Read column MobileandAccessories
   ' Print "Row " & ": " & vSlNo & " :- " & vMobileandAccessories
    Dim vResult
    vResult = vSlNo & " :- " & vMobileandAccessories
    Print "result" & vResult
   'To update the Notepad 
    Call fnWriteNotepadwithResults(i, vResult)
Next

'Save
objWorkbook.Save
objWorkbook.Close
objExcel.Quit

Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing


End Function



Public Function fnWriteNotepadwithResults(i, vResult)
' Create a FSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

' path
filePath = "C:\Users\MBaladhandapani\Desktop\UFT_Eval_PS\Notepad_MobileandAccessories_Result.txt"

' Create or Open the file for writing (Mode 8 - Append)
Set objFile = objFSO.OpenTextFile(filePath, 8, True)

' Write new data into the file
objFile.WriteLine vResult & vbCrLf 
Print vResult
		
' Close the file
objFile.Close

' Cleanup
Set objFile = Nothing
Set objFSO = Nothing
	
End Function


