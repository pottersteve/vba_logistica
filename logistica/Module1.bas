Attribute VB_Name = "Module1"
Dim driver As WebDriver
Dim n As Integer


'LIKE THE MAIN FUNCTION!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Sub RectangleRoundedCorners1_Click()
    Const URL As String = "https://wiz.elj-automotive.ro/"
    Set driver = New EdgeDriver
    n = 1
    
    With driver
            .SetCapability "ms:edgeOptions", "{""args"":[""--headless""]}"
            .Start "Edge"
            .Get URL
    End With
    
    Call beginning
    Call macro_raport1
    Call macro_raport2

    MsgBox "Rapoarte introduse"
    
    'Update
    ThisWorkbook.RefreshAll
    ThisWorkbook.Save
    
    Application.Wait Now + TimeValue("00:00:05")
    driver.Quit
    Set driver = Nothing
    
    Call getFromPivotAndMakeFinalShit
    
    'save and shit
    ThisWorkbook.Save
    
End Sub



Sub beginning() 'aici incepe priority

    driver.Wait 5000

    Dim username As String
    Dim password As String
    Dim path As String
    Dim email As String
    Dim subject As String
    
    username = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D2").Value
    password = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D3").Value
    path = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D4").Value
    
    Dim dropdown As Selenium.WebElement
    Dim okButton As Selenium.WebElement
    Dim descarcat As Boolean
    Dim email_sent As Boolean
    descarcat = False
    

    'LOGARE'''''''''''''''''''''''''''''''''''
    driver.FindElementByCss("input[placeholder='User Name']", 10000).SendKeys (username)
    driver.FindElementByCss("input[placeholder='Password']", 10000).SendKeys (password)
    driver.FindElementByCss("button.loginlabel_buttonStyle1", 100000).Click
    

    'error handling here pentru 1/20 de cazuri in care da eroare
    'NO ERROR HANDLING, MAKE IT
    OnErrorSub
    
    
    'limba i guess (romana e 3, dar punem 2 pentru ca indexarea incepe de la 0)
    'Set dropdown = driver.FindElementByClass("ListBox_Language")
    'dropdown.AsSelect.SelectByIndex 2
    'Set okButton = driver.FindElementByXPath("//div[text()='OK']")
    'okButton.Click
    
    'x la limba
    driver.FindElementByClass("xbutton-up").Click
    
    
    'ALEGE COMPANIA''''''''''''''''''''''''
    'CLICK PENTRU LISTA DE OPTIUNI
    '<div class="DropdownSelectTrigger-module_inputWrapper__jbjku DropdownSelectTrigger-module_box__cS8R5 DropdownSelectTrigger-module_medium__WZhcc DropdownSelectTrigger-module_inputWrapperPaddingAll__ioOJj"><div class="DropdownSelectTrigger-module_input__-lIM6"><span class="DropdownSelectTrigger-module_placeholder__GaogO">International demo</span></div><div class="Flex-module_component__bBMUM Flex-module_row__f-hHa Flex-module_align-center__l1xmu Flex-module_justify-space-between__dWgE4 Flex-module_inline__bTjgU"> <div class="Icon-module_component__22QhN icon-new_arrow-bottom text-color-primary-700 Icon-module_sizeDefault__9k9wJ" aria-label="new_arrow-bottom"></div> </div></div>
    driver.FindElementByCss("div.DropdownSelectTrigger-module_inputWrapper__jbjku").Click
    'ALEGE DIN LISTA
    '<div class="pui-dropdown-item DropdownItem-module_item__5DI-a DropdownItem-module_itemBorder__-124u" data-test-id="dropdown-list-item"><div class="text DropdownItem-module_ellipsis__peR58 text-paragraph-2 text-color-primary-700 text-line-breaks" data-test-id="common-text">ELJ Automotive</div><div class="text DropdownItem-module_defaultDesc__fkpX9 DropdownItem-module_ellipsis__peR58 text-paragraph-2 text-color-secondary-400 text-line-breaks" data-test-id="common-text"></div></div>
    driver.FindElementByCss("div.pui-dropdown-item.DropdownItem-module_item__5DI-a.DropdownItem-module_itemBorder__-124u:nth-of-type(4)").Click
    'CLICK OK
    driver.FindElementByCss("button.Button-module_component__yrLFG.Button-module_primary__3i9Qe.Button-module_contentWidth__fT0IG.Button-module_medium__SSp73.undefined.Button-module_roundedCorners__g1q5q.Button-module_wideButton__V03Mq").Click
    '<button type="button" tabindex="0" class="Button-module_component__yrLFG Button-module_primary__3i9Qe Button-module_contentWidth__fT0IG Button-module_medium__SSp73 undefined Button-module_roundedCorners__g1q5q Button-module_wideButton__V03Mq" data-test-id="common-button"><span class="Button-module_buttonContent__xJneS"><span class="Button-module_textWrap__9vFmu">OK</span></span></button>

End Sub



Sub macro_raport1()
    'DROPDOWS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'driver.FindElementByCss("a#merk_menu_item_58").Click
    Dim targetElement1 As WebElement
    Dim targetElement2 As WebElement
    Dim targetElement3 As WebElement
    Dim targetElement4 As WebElement
    Dim targetElement5 As WebElement
    Dim jsScript As String
    
    'Create the exact JavaScript to simulate a real human mouse click (PENTRU REACT)
    jsScript = "arguments[0].dispatchEvent(new MouseEvent('click', {" & _
               "bubbles: true, " & _
               "cancelable: true, " & _
               "view: window" & _
               "}));"
               
    driver.Wait 500
    
    'nav
    Set targetElement1 = driver.FindElementById("merk_menu_item_59", 10000)
    driver.ExecuteScript jsScript, targetElement1
    
    driver.Wait 500
    
    'meniu 1
    Set targetElement2 = driver.FindElementByCss("ul.dropMenu.visible > li > a#merk_menu_item_6012", 100000)
    driver.ExecuteScript jsScript, targetElement2
    
    driver.Wait 500
    
    'meniu 2
    Set targetElement3 = driver.FindElementById("merk_menu_item_1987", 10000)
    driver.ExecuteScript jsScript, targetElement3
    
    driver.Wait 500
    
    'meniu 3
    Set targetElement3 = driver.FindElementById("merk_menu_item_295", 10000)
    driver.ExecuteScript jsScript, targetElement3
    
    driver.Wait 500
    
    'meniu 4
    Set targetElement4 = driver.FindElementById("merk_menu_item_6267", 10000)
    driver.ExecuteScript jsScript, targetElement4
    
    driver.Wait 500
    
    'meniu 5
    Set targetElement5 = driver.FindElementById("merk_menu_item_4706", 10000)
    driver.ExecuteScript jsScript, targetElement5
    
    driver.Wait 500
    
    
    
    'MENIURI!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    'buton spreadsheet
    driver.FindElementByCss("button.actionCard-module_typeCard_wjjYM.actionCard-module_selected_sDuP5", 10000).Click
    driver.Wait 500
    'buton ok 1
    driver.FindElementByCss("button.Button-module_component__yrLFG.Button-module_primary__3i9Qe.Button-module_regularFixed__4w9eE", 100000).Click
    driver.Wait 1000
    
    'date
    '<input type="text" class="inputDialogTextBoxStyle inputfield_1" spellcheck="false" maxlength="40" dir="ltr" style="width: 180px; direction: ltr; text-align: left;">
    'driver.FindElementByCss("input.inputDialogTextBoxStyle.inputfield_1:nth-of-type(1)").Clear.SendKeys ("start of this month")
    '<input type="text" class="inputDialogTextBoxStyle inputfield_1" spellcheck="false" maxlength="40" dir="ltr" style="width: 180px; direction: ltr;">
    'driver.FindElementByCss("input.inputDialogTextBoxStyle.inputfield_1:nth-of-type(2)").Clear.SendKeys ("end of this month")
    
    'preia date
    Dim data_first As String
    Dim data_second As String
    
    data_first = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D11").Value
    data_second = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D12").Value
    'MsgBox data_first
    
    'schimba Data
    ThisWorkbook.Worksheets("final").Range("A1").MergeArea.Value = data_first
    
    
    ' Note the backslashes to force literal slashes regardless of local Windows settings???????????????????
    'data_first = Format(data_first, "dd" & "/" & "mm" & "/" & "yy")
    'data_second = Format(data_second, "dd" & "/" & "mm" & "/" & "yy")
    data_first = Format(data_first, "mm\/dd\/yy")
    data_second = Format(data_second, "mm\/dd\/yy")
    'MsgBox data_first
    
    
    Set input1 = driver.FindElementByXPath("(//input[contains(@class, 'inputfield_1')])[1]")
    Set input2 = driver.FindElementByXPath("(//input[contains(@class, 'inputfield_1')])[2]")
    
    
    jsReactSetter = "var element = arguments[0]; " & _
                    "var val = '" & data_first & "'; " & _
                    "var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set; " & _
                    "if (setter) { setter.call(element, val); } else { element.value = val; } " & _
                    "element.dispatchEvent(new Event('input', { bubbles: true })); " & _
                    "element.dispatchEvent(new Event('change', { bubbles: true }));"
    
    driver.ExecuteScript jsReactSetter, input1
    
    
    
    jsReactSetter = "var element = arguments[0]; " & _
                    "var val = '" & data_second & "'; " & _
                    "var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set; " & _
                    "if (setter) { setter.call(element, val); } else { element.value = val; } " & _
                    "element.dispatchEvent(new Event('input', { bubbles: true })); " & _
                    "element.dispatchEvent(new Event('change', { bubbles: true }));"
    
    driver.ExecuteScript jsReactSetter, input2
    
    
    
    'buton ok 2
    '<div tabindex="0" class="gwt-PushButton MiddleDialogButton inlineManual_inputDlg_ok gwt-PushButton-up gwt-PushButton-focused" role="button" aria-pressed="false" style="text-align: center;"><input type="text" tabindex="-1" role="presentation" style="opacity: 0; height: 1px; width: 1px; z-index: -1; overflow: hidden; position: absolute;"><div class="html-face">OK</div></div>
    Dim lastOK As String
    lastOK = "div.MiddleDialogButton.inlineManual_inputDlg_ok"
    If Not IsEmpty(driver.FindElementByCss(lastOK)) Then
        driver.FindElementByCss(lastOK).Click
        descarcat = True
    End If
    
    Dim targetInput1 As Object
    Dim jsScript1 As String
    
    Set targetInput1 = driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[1]")
    jsScript1 = "var input = arguments[0];" & _
               "var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;" & _
               "nativeInputValueSetter.call(input, '*');" & _
               "var event = new Event('input', { bubbles: true });" & _
               "input.dispatchEvent(event);"
    driver.ExecuteScript jsScript1, targetInput1
    
    'driver.Wait 1000
    
    Dim targetInput2 As Object
    Dim jsScript2 As String
    
    Set targetInput2 = driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[3]")
    jsScript2 = "var input = arguments[0];" & _
               "var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;" & _
               "nativeInputValueSetter.call(input, '*');" & _
               "var event = new Event('input', { bubbles: true });" & _
               "input.dispatchEvent(event);"
    driver.ExecuteScript jsScript2, targetInput2
    
    'driver.Wait 1000
    
    '<input dir="DEFAULT" class="Input-module_input__0SCRX" placeholder="" value="">
    Dim targetInput3 As Object
    Dim jsScript3 As String
    
    Set targetInput3 = driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[5]")
    jsScript3 = "var input = arguments[0];" & _
               "var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;" & _
               "nativeInputValueSetter.call(input, '3*');" & _
               "var event = new Event('input', { bubbles: true });" & _
               "input.dispatchEvent(event);"
    driver.ExecuteScript jsScript3, targetInput3
    
    'driver.Wait 1000
    
    'buton 3
    '<div tabindex="0" class="gwt-PushButton MiddleDialogButton gwt-PushButton-up inlineManual_inputDlg_ok gwt-PushButton-focused" role="button" aria-pressed="false" style="text-align: center;"><input type="text" tabindex="-1" role="presentation" style="opacity: 0; height: 1px; width: 1px; z-index: -1; overflow: hidden; position: absolute;"><div class="html-face">OK</div></div>
    Dim lastlastOK As String
    Dim element As Object
    Dim actions As Object
    
    lastlastOK = "div.MiddleDialogButton.inlineManual_inputDlg_ok"
    
    ' Check if the element exists by counting elements to avoid runtime crashes
    If driver.FindElementsByCss(lastlastOK).Count > 0 Then
        
        'Get the specific element
        Set element = driver.FindElementByCss(lastlastOK)
        
        'Initialize the Actions class to simulate mouse movement
        Set actions = driver.actions
        
        'Move the mouse cursor to the center of the element, then click
        actions.MoveToElement(element).Click.Perform
        
        descarcat = True
    End If
    
    driver.Wait 10000
    
    Dim array_sales_order As Variant
    array_sales_order = Array("Order", "Part Number", "Quantity", "Balance")
    Call OpenAndExtractData_1(array_sales_order)
    
End Sub


Sub OpenAndExtractData_1(array_column_names As Variant)
    Dim path As String
    Dim file_name As String
    Dim mainWB As Workbook
    Dim wsMaster As Worksheet
    Dim sourceWB As Workbook
    Dim sourceWS As Worksheet
    Dim newWS As Worksheet
    Dim lastRow As Long
    Dim path_pentru_fisierul_curent As String
    path_pentru_fisierul_curent = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D5").Value
    
    Application.ScreenUpdating = False 'make the macro run faster and look smoother
    
    Set mainWB = Workbooks.Open(path_pentru_fisierul_curent & "macro logistica nou.xlsm")
    'MsgBox "am gasit mainWB: " + mainWB.Name
    Set wsMaster = mainWB.Worksheets("Rapoarte")
    'MsgBox "am gasit rapoarte"
    
    file_name = Rename()
    path = wsMaster.Range("D4").Value
    If Right(path, 1) <> "\" Then
        path = path & "\"
    End If
    'MsgBox "am gasit path ul"
    
    Set sourceWB = Workbooks.Open(path & file_name)
    'MsgBox "am gasit sourceWB: " + sourceWB.Name
    Set sourceWS = sourceWB.Worksheets("Datasheet")
    
    'On Error Resume Next
    'On Error GoTo 0
    Set newWS = mainWB.Sheets("sales_order")
    'If newWS Is Nothing Then
        'With mainWB
            'Set newWS = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            'newWS.Name = new_worksheet
        'End With
        'MsgBox "am adaugat worksheet not"
    'End If
    
    If newWS Is Nothing Then
        MsgBox "nu merge"
        Exit Sub
    End If
    
    Dim lenght As Integer
    Dim col As Long
    Dim colLetter As String
    col = 1
    newWS.Cells.Clear
    Do While sourceWS.Cells(1, col).Value <> ""
    
        ' Copy the entire column directly, no matching required
        sourceWS.Columns(col).Copy Destination:=newWS.Columns(n)
        
        ' Increment the destination column pointer
        n = n + 1
        
        ' Move to the next source column
        col = col + 1
    Loop
        

    
    'sourceWS.Columns(3).Copy Destination:=newWS.Columns(n)
    'n = n + 1
    'sourceWS.Columns(5).Copy Destination:=newWS.Columns(n)
    'n = n + 1
    'sourceWS.Columns(7).Copy Destination:=newWS.Columns(n)
    'n = n + 1
    'sourceWS.Columns(8).Copy Destination:=newWS.Columns(n)
    'n = n + 1
    
    'Clean up and close
    Application.CutCopyMode = False
    
    ' Close the source file (we don't need to save changes to it)
    sourceWB.Close SaveChanges:=False
    
    ' Save and close the Master workbook since we just pasted new data into it
    'mainWB.Close SaveChanges:=True
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    n = 1
End Sub



Sub macro_raport2()
    
    Dim dropdown As Selenium.WebElement
    Dim okButton As Selenium.WebElement
    Dim descarcat As Boolean
    Dim email_sent As Boolean
    descarcat = False
    
  
    'DROPDOWS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'driver.FindElementByCss("a#merk_menu_item_58").Click
    Dim targetElement1 As WebElement
    Dim targetElement2 As WebElement
    Dim targetElement3 As WebElement
    Dim jsScript As String
    
    'Create the exact JavaScript to simulate a real human mouse click (PENTRU REACT)
    jsScript = "arguments[0].dispatchEvent(new MouseEvent('click', {" & _
               "bubbles: true, " & _
               "cancelable: true, " & _
               "view: window" & _
               "}));"
               
    driver.Wait 500
    
    'nav
    Set targetElement1 = driver.FindElementById("merk_menu_item_1201", 10000)
    driver.ExecuteScript jsScript, targetElement1
    
    driver.Wait 500
    
    'meniu 1
    Set targetElement2 = driver.FindElementByCss("ul.dropMenu.visible > li > a#merk_menu_item_1315", 100000)
    driver.ExecuteScript jsScript, targetElement2
    
    driver.Wait 500
    
    'meniu 2
    Set targetElement3 = driver.FindElementById("merk_menu_item_4537", 10000)
    driver.ExecuteScript jsScript, targetElement3
    
    driver.Wait 500
    
    'meniu 3
    Set targetElement3 = driver.FindElementById("merk_menu_item_1323", 10000)
    driver.ExecuteScript jsScript, targetElement3
    
    driver.Wait 500
    
    
    
    'MENIURI!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    'buton spreadsheet
    driver.FindElementByCss("button.actionCard-module_typeCard_wjjYM.actionCard-module_selected_sDuP5", 10000).Click
    driver.Wait 500
    'buton ok 1
    driver.FindElementByCss("button.Button-module_component__yrLFG.Button-module_primary__3i9Qe.Button-module_regularFixed__4w9eE", 100000).Click
    driver.Wait 1000
    
    'date
    '<input type="text" class="inputDialogTextBoxStyle inputfield_1" spellcheck="false" maxlength="40" dir="ltr" style="width: 180px; direction: ltr; text-align: left;">
    'driver.FindElementByCss("input.inputDialogTextBoxStyle.inputfield_1:nth-of-type(1)").Clear.SendKeys ("start of this month")
    '<input type="text" class="inputDialogTextBoxStyle inputfield_1" spellcheck="false" maxlength="40" dir="ltr" style="width: 180px; direction: ltr;">
    'driver.FindElementByCss("input.inputDialogTextBoxStyle.inputfield_1:nth-of-type(2)").Clear.SendKeys ("end of this month")
    
    'SCHIMBA PARAMETRII ROMANA/ENGLEZA
    
    
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '<div class="Input-module_inputWrapper__a4cDW Input-module_small__Xs1iD Input-module_box__CtS3D Input-module_inputWrapperPaddingStart__MVRl1"><input dir="DEFAULT" class="Input-module_input__0SCRX Input-module_forceLTR__PkfwU" placeholder="" value="*"><div class="Flex-module_component__bBMUM Flex-module_row__f-hHa Flex-module_align-center__l1xmu Flex-module_justify-space-between__dWgE4 Flex-module_inline__bTjgU"> <div class="Icon-module_component__22QhN icon-new_iconSearch ProcSearch-module_searchIcon_RXcdF Icon-module_standart__Qj34C Icon-module_sizeDefault__9k9wJ Icon-module_hoverable__zDw4p" aria-label="new_iconSearch"></div><div class="Icon-module_component__22QhN icon-icon-arrow_drop_down Icon-module_standart__Qj34C Icon-module_sizeDefault__9k9wJ Icon-module_hoverable__zDw4p" aria-label="icon-arrow_drop_down"></div> </div></div>
    '<input dir="auto" class="Input-module_input__0SCRX Input-module_forceLTR__PkfwU" placeholder="" value="DCPF">
    'driver.FindElementByCss("div.priModalDialog > div > table > tbody > td.popupMiddleCenter > div > div > table >tbody > tr:nth-of-type(2) > td > table > tbody > tr > td > table > tr:nth-of-type(1) > td:nth-of-type(2) > table > tbody > tr > td:nth-of-type(2) > input.Input-module_input__0SCRX").Clear
    'driver.FindElementByCss("div.priModalDialog > div > table > tbody > td.popupMiddleCenter > div > div > table >tbody > tr:nth-of-type(2) > td > table > tbody > tr > td > table > tr:nth-of-type(1) > td:nth-of-type(2) > table > tbody > tr > td:nth-of-type(2) > div > div > div > div > div > div > div > input").SendKeys "DPF"
    'driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[1]").Clear

    'driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[1]").SendKeys (Keys.chord(Keys.Control, "a"))
    'driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[1]").SendKeys ("vittu")
    
    'Dim target_input
    'Set target_input = driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[1]")
    'driver.ExecuteScript "arguments[0].setAttribute('placeholder', 'vittu');", target_input
    'driver.ExecuteScript "arguments[0].setAttribute('value', 'vittu');", target_input
    
    Dim targetInput1 As Object
    Dim jsScript1 As String
    
    Set targetInput1 = driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[1]")
    jsScript1 = "var input = arguments[0];" & _
               "var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;" & _
               "nativeInputValueSetter.call(input, 'DPF');" & _
               "var event = new Event('input', { bubbles: true });" & _
               "input.dispatchEvent(event);"
    driver.ExecuteScript jsScript1, targetInput1
    
    
    Dim targetInput2 As Object
    Dim jsScript2 As String
    
    Set targetInput2 = driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[3]")
    jsScript2 = "var input = arguments[0];" & _
               "var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;" & _
               "nativeInputValueSetter.call(input, 'D*');" & _
               "var event = new Event('input', { bubbles: true });" & _
               "input.dispatchEvent(event);"
    driver.ExecuteScript jsScript2, targetInput2
    
    
    Dim targetInput3 As Object
    Dim jsScript3 As String
    
    Set targetInput3 = driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[5]")
    jsScript3 = "var input = arguments[0];" & _
               "var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;" & _
               "nativeInputValueSetter.call(input, '3*');" & _
               "var event = new Event('input', { bubbles: true });" & _
               "input.dispatchEvent(event);"
    driver.ExecuteScript jsScript3, targetInput3
    
    'driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[1]").SendKeys "vittu"
    'Application.SendKeys (Keys.Tab)
    'Application.SendKeys (Keys.Tab)
    'Application.SendKeys "D*"
    'Application.SendKeys (Keys.Tab)
    'Application.SendKeys (Keys.Tab)
    'Application.SendKeys "3*"
    
    'driver.FindElementByXPath("//div[text()='Locatie']/ancestor::tr[1]//div[not(contains(@style, 'display: none'))]/input").Clear
    'driver.FindElementByXPath("//div[text()='Locatie']/ancestor::tr[1]//div[not(contains(@style, 'display: none'))]/input").SendKeys "DPF"
    'driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[3]").Clear
    'driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[3]").SendKeys (Keys.chord(Keys.Control, "a", Keys.Delete))
    'driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[3]").SendKeys "D*"
    
    'driver.FindElementByXPath("(//input[contains(@class, 'Input-module_input__0SCRX')])[4]").SendKeys "3*"
    
    'buton ok 2
    '<div tabindex="0" class="gwt-PushButton MiddleDialogButton inlineManual_inputDlg_ok gwt-PushButton-up gwt-PushButton-focused" role="button" aria-pressed="false" style="text-align: center;"><input type="text" tabindex="-1" role="presentation" style="opacity: 0; height: 1px; width: 1px; z-index: -1; overflow: hidden; position: absolute;"><div class="html-face">OK</div></div>
    Dim lastOK As String
    lastOK = "div.MiddleDialogButton.inlineManual_inputDlg_ok"
    If Not IsEmpty(driver.FindElementByCss(lastOK)) Then
        driver.FindElementByCss(lastOK).Click
        descarcat = True
    End If
    
    driver.Wait 1000
    

    Dim array_current_inventory As Variant
    array_current_inventory = Array("Part Number", "Balance")
    Call OpenAndExtractData_2(array_current_inventory)
End Sub



Sub OpenAndExtractData_2(array_column_names As Variant)
    Dim path As String
    Dim file_name As String
    Dim mainWB As Workbook
    Dim wsMaster As Worksheet
    Dim sourceWB As Workbook
    Dim sourceWS As Worksheet
    Dim newWS As Worksheet
    Dim lastRow As Long
    Dim path_pentru_fisierul_curent As String
    path_pentru_fisierul_curent = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D5").Value
    
    Application.ScreenUpdating = False 'make the macro run faster and look smoother
    
    Set mainWB = Workbooks.Open(path_pentru_fisierul_curent & "macro logistica nou.xlsm")
    'MsgBox "am gasit mainWB: " + mainWB.Name
    Set wsMaster = mainWB.Worksheets("Rapoarte")
    'MsgBox "am gasit rapoarte"

    
    driver.Wait 5000
    
    file_name = Rename()
    path = wsMaster.Range("D4").Value
    If Right(path, 1) <> "\" Then
        path = path & "\"
    End If
    'MsgBox "am gasit path ul"
    
    Set sourceWB = Workbooks.Open(path & file_name)
    'MsgBox "am gasit sourceWB: " + sourceWB.Name
    Set sourceWS = sourceWB.Worksheets("Datasheet")
    
    'On Error Resume Next
    'On Error GoTo 0
    Set newWS = mainWB.Sheets("current_inventory")
    'If newWS Is Nothing Then
        'With mainWB
            'Set newWS = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            'newWS.Name = new_worksheet
        'End With
        'MsgBox "am adaugat worksheet not"
    'End If
    
    If newWS Is Nothing Then
        MsgBox "nu merge"
        Exit Sub
    End If
    
    Dim lenght As Integer
    Dim col As Long
    Dim colLetter As String
    col = 1
    newWS.Cells.Clear
    Do While sourceWS.Cells(1, col).Value <> ""
    
        ' Copy the entire column directly, no matching required
        sourceWS.Columns(col).Copy Destination:=newWS.Columns(n)
        
        ' Increment the destination column pointer
        n = n + 1
        
        ' Move to the next source column
        col = col + 1
    Loop
        

    
    'sourceWS.Columns(3).Copy Destination:=newWS.Columns(n)
    'n = n + 1
    'sourceWS.Columns(5).Copy Destination:=newWS.Columns(n)
    'n = n + 1
    'sourceWS.Columns(7).Copy Destination:=newWS.Columns(n)
    'n = n + 1
    'sourceWS.Columns(8).Copy Destination:=newWS.Columns(n)
    'n = n + 1
    
    'Clean up and close
    Application.CutCopyMode = False
    
    ' Close the source file (we don't need to save changes to it)
    sourceWB.Close SaveChanges:=False
    
    ' Save and close the Master workbook since we just pasted new data into it
    'mainWB.Close SaveChanges:=True
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    n = 1
End Sub



Sub foldere()
    Dim file_name As String
    Dim BeforeFirstBlankCell_C As Long
    Dim BeforeFirstBlankCell_J As Long
    Dim BeforeFirstBlankCell_K As Long
    Dim sum_j As Double
    Dim sum_k As Double
    Dim nr_zile As Range
    Dim nr_zile_unice As Integer
    Dim opened_file As Workbook
    Dim vanzari_nete As Double
    
    file_name = Rename()
    
    'open file and shit
    driver.Wait 5000
    'FILE NAME AHH EXTENSIONS AHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHH
    If Right((file_name), 5) <> ".xlsx" Then
        If Right((file_name), 4) = ".xls" Then
            file_name = file_name & "x"
        Else
            file_name = file_name & ".xlsx"
        End If
    End If
    
    MsgBox path + file_name
    Set opened_file = Workbooks.Open(Filename:=path & file_name)
    
    'sums and shit
    BeforeFirstBlankCell_J = Workbooks(file_name).Worksheets("DataSheet").Range("J2").End(xlDown).Row
    BeforeFirstBlankCell_K = Workbooks(file_name).Worksheets("DataSheet").Range("K2").End(xlDown).Row
    sum_j = Application.WorksheetFunction.Sum(opened_file.Worksheets("DataSheet").Range("J2:J" & BeforeFirstBlankCell_J))
    sum_k = Application.WorksheetFunction.Sum(opened_file.Worksheets("DataSheet").Range("K2:K" & BeforeFirstBlankCell_K))
    
    vanzari_nete = sum_k - sum_j
    vanzari_nete = Round(vanzari_nete, 2)
    
    BeforeFirstBlankCell_C = Workbooks(file_name).Worksheets("DataSheet").Range("C2").End(xlDown).Row
    Set nr_zile = Workbooks(file_name).Worksheets("DataSheet").Range("C2:C" & BeforeFirstBlankCell_C)
    nr_zile_unice = Application.WorksheetFunction.Count(Application.WorksheetFunction.Unique(nr_zile))
    

    Workbooks(file_name).Close SaveChanges:=False

    'pick up the file and send outlook email
    Call SendMailWAttachement(file_name, vanzari_nete, nr_zile_unice)
    driver.Wait 5000
    
    'delete file
    Kill (path + file_name)
    
    
    'cod tibi??
    'Dim folderasp As String
    'folderasp = "D:\Ion Tiberiu\Desktop\01. PRELUCRARI\prelucrareFILES\DB" ' Replace with your folder path
    'Shell "explorer.exe " & folderasp, vbNormalFocus
    'ActiveWorkbook.FollowHyperlink Address:="D:\Ion Tiberiu\Downloads", NewWindow:=True

End Sub



Public Function Rename() As String 'doesnt rename, but picks up the name. and checks that it was generated the last minute
    Dim path As String
    path = Workbooks("macro logistica nou.xlsm").Worksheets("Rapoarte").Range("D4").Value

    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim strName As String
    Dim varDate As Variant
    Dim fileAgeInSeconds As Double
    Dim isFileReady As Boolean
    
    'Microsoft Scripting runtime
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    isFileReady = False
    
    'LOOP until a file created/modified less than 1 minute ago is found
    Do While True
        Set objFolder = objFSO.GetFolder(path)
        varDate = 0
        strName = ""
        
        'CHECK DATE
        For Each objFile In objFolder.Files
            If objFile.DateLastModified > varDate Then
                varDate = objFile.DateLastModified
                strName = objFile.Name
            End If
        Next objFile
        
        'CHECK AGE
        If strName <> "" Then
            'Calculate the difference in seconds between Now and the file's last modified date
            fileAgeInSeconds = DateDiff("s", varDate, Now)
            
            ' !!!!THE IMPORTANT ONE!!!! check if it was modified less than 60 seconds ago
            If fileAgeInSeconds >= 0 And fileAgeInSeconds < 60 Then
                Exit Do 'exit the wait loop because latest is good enough
            End If
        End If
        
        'wait 5 seconds before checking the folder again
        Application.Wait (Now + TimeValue("00:00:03"))
        DoEvents ' "Keeps Excel responsive during the wait" - gemini
    Loop
    
    'MsgBox strName & " - is latest file - " & varDate
    
    clean_name = objFSO.GetBaseName(strName) 'cleans up and erases the extensions
    
    Rename = clean_name
    
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set objFile = Nothing
End Function



Sub getFromPivotAndMakeFinalShit()
    Dim i As Integer: i = 1
    Dim gasit As Boolean
    Dim finalSheet As Worksheet
    Set finalSheet = ThisWorkbook.Worksheets("final")
    
    ' Clear old data
    finalSheet.Rows(5 & ":" & finalSheet.Rows.Count).Delete
    
    Dim pivot_table_sales As PivotTable
    Dim pivot_field_sales As PivotField
    Dim codes_sales As Range
    
    Set pivot_table_sales = ThisWorkbook.Worksheets("sales_order_pivot").PivotTables("PivotTable1")
    Set pivot_field_sales = pivot_table_sales.RowFields(1)
    Set codes_sales = pivot_field_sales.DataRange
    
    Dim pivot_table_inventory As PivotTable
    Dim pivot_field_inventory As PivotField
    Dim codes_inventory As Range
    
    Set pivot_table_inventory = ThisWorkbook.Worksheets("current_inventory_pivot").PivotTables("PivotTable2")
    Set pivot_field_inventory = pivot_table_inventory.RowFields(1)
    Set codes_inventory = pivot_field_inventory.DataRange
    
    ' Safety check in case a pivot table is empty
    If codes_sales Is Nothing Or codes_inventory Is Nothing Then
        MsgBox "One of the pivot tables has no data!", vbCritical
        Exit Sub
    End If
    
    ' COMPARING
    Dim code As Range, code2 As Range
    Dim nextRow As Long
    
    ThisWorkbook.Unprotect
    
    ' Speed up execution by turning off screen updates
    Application.ScreenUpdating = False
    
    For Each code In codes_sales
        gasit = False
        
        ' Skip empty cells or total rows if they sneak in
        If code.Offset(0, -1).Value <> "" And Not InStr(1, code.Offset(0, -1).Value, "(blank)", vbTextCompare) > 0 Then
            
            For Each code2 In codes_inventory
                
                'MsgBox "Sales Code: [" & code.Value & "] Length: " & Len(code.Value) & vbCrLf & "Inv Code: [" & code2.Value & "] Length: " & Len(code2.Value)

                
                If StrComp(Trim(CStr(code.Offset(0, -1).Value)), Trim(CStr(code2.Value)), vbTextCompare) = 0 Then
                    gasit = True
                   
                    nextRow = finalSheet.Cells(finalSheet.Rows.Count, "A").End(xlUp).Row + 1
                    
                    
                    
                    If code.Offset(0, 2).Value > code2.Offset(0, 1).Value Then
                        finalSheet.Cells(nextRow, 1).Value = i ' Column A: Number
                        finalSheet.Cells(nextRow, 2).Value = code.Offset(0, 0).Value ' Column B: Client
                        finalSheet.Cells(nextRow, 3).Value = code.Offset(0, -1).Value ' Column C: Code
                        finalSheet.Cells(nextRow, 4).Value = code.Offset(0, 1).Value ' Column D: Description
                        finalSheet.Cells(nextRow, 5).Value = code.Offset(0, 2).Value ' Column E: Comanda
                        finalSheet.Cells(nextRow, 6).Value = code2.Offset(0, 1).Value ' Column G: Inventory Value
                        finalSheet.Cells(nextRow, 7).Value = Abs(finalSheet.Cells(nextRow, 6).Value - finalSheet.Cells(nextRow, 5).Value) ' Column H: Total
                        i = i + 1
                    Else
                        finalSheet.Cells(nextRow, 7).Value = Empty
                    End If
                    
                    
                    Exit For ' Match found, stop checking inventory for this code
                End If
            Next code2
        
            ' If no match was found in the inner loop
            If gasit = False Then
                nextRow = finalSheet.Cells(finalSheet.Rows.Count, "A").End(xlUp).Row + 1
                    
                finalSheet.Cells(nextRow, 1).Value = i ' Column A: Number
                finalSheet.Cells(nextRow, 2).Value = code.Offset(0, 0).Value ' Column B: Client
                finalSheet.Cells(nextRow, 3).Value = code.Offset(0, -1).Value ' Column C: Code
                finalSheet.Cells(nextRow, 4).Value = code.Offset(0, 1).Value ' Column D: Description
                finalSheet.Cells(nextRow, 5).Value = code.Offset(0, 2).Value ' Column E: Comanda
                finalSheet.Cells(nextRow, 6).Value = 0 ' Column G: Inventory Value (0 because not found)
                finalSheet.Cells(nextRow, 7).Value = Abs(finalSheet.Cells(nextRow, 6).Value - finalSheet.Cells(nextRow, 5).Value) ' Column H: Total
                
                i = i + 1
            End If
        End If
    Next code
    
    Application.ScreenUpdating = True
    MsgBox "Table updated successfully!"
End Sub



Sub OnErrorSub()
    On Error GoTo EroareLimbi
        driver.FindElementByCss("button.Button-module_component__yrLFG.Messages-module_footerBtn_N7e8F", 10000).Click
    
Done:
    Exit Sub
EroareLimbi: limbajprost
            
End Sub


Sub limbajprost()
MsgBox "Eroare din Priority. Incearca din nou"
            driver.Quit
            Set driver = Nothing
End Sub


