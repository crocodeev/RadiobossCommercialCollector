﻿Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$main                            = New-Object system.Windows.Forms.Form
$main.ClientSize                 = New-Object System.Drawing.Point(622,266)
$main.text                       = "commercialCollector"
$main.TopMost                    = $false

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 454
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(150,17)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$GroupboxPrefixes                = New-Object system.Windows.Forms.Groupbox
$GroupboxPrefixes.height         = 96
$GroupboxPrefixes.width          = 407
$GroupboxPrefixes.text           = "Префиксы"
$GroupboxPrefixes.location       = New-Object System.Drawing.Point(18,52)

$ReklamaCheckBox                 = New-Object system.Windows.Forms.CheckBox
$ReklamaCheckBox.text            = "Реклама/Reklama"
$ReklamaCheckBox.AutoSize        = $false
$ReklamaCheckBox.width           = 140
$ReklamaCheckBox.height          = 18
$ReklamaCheckBox.location        = New-Object System.Drawing.Point(9,23)
$ReklamaCheckBox.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$JingleCheckBox                  = New-Object system.Windows.Forms.CheckBox
$JingleCheckBox.text             = "Джингл/Jingle"
$JingleCheckBox.AutoSize         = $false
$JingleCheckBox.width            = 113
$JingleCheckBox.height           = 18
$JingleCheckBox.location         = New-Object System.Drawing.Point(149,23)
$JingleCheckBox.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BilboardCheckBox                = New-Object system.Windows.Forms.CheckBox
$BilboardCheckBox.text           = "Билборд/Billboard"
$BilboardCheckBox.AutoSize       = $false
$BilboardCheckBox.width          = 143
$BilboardCheckBox.height         = 18
$BilboardCheckBox.location       = New-Object System.Drawing.Point(262,23)
$BilboardCheckBox.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$CustomCheckBox                  = New-Object system.Windows.Forms.CheckBox
$CustomCheckBox.text             = "Пользовательский префикс"
$CustomCheckBox.AutoSize         = $false
$CustomCheckBox.width            = 150
$CustomCheckBox.height           = 18
$CustomCheckBox.location         = New-Object System.Drawing.Point(9,53)
$CustomCheckBox.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.width                  = 181
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(199,50)
$TextBox2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$GroupboxOptions                 = New-Object system.Windows.Forms.Groupbox
$GroupboxOptions.height          = 96
$GroupboxOptions.width           = 171
$GroupboxOptions.text            = "Включая размещённые"
$GroupboxOptions.location        = New-Object System.Drawing.Point(433,52)

$CheckBoxPlaylist                = New-Object system.Windows.Forms.CheckBox
$CheckBoxPlaylist.text           = "В плейлистах"
$CheckBoxPlaylist.AutoSize       = $false
$CheckBoxPlaylist.width          = 120
$CheckBoxPlaylist.height         = 20
$CheckBoxPlaylist.location       = New-Object System.Drawing.Point(10,28)
$CheckBoxPlaylist.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$CheckBoxFolders                 = New-Object system.Windows.Forms.CheckBox
$CheckBoxFolders.text            = "В папках"
$CheckBoxFolders.AutoSize        = $false
$CheckBoxFolders.width           = 95
$CheckBoxFolders.height          = 20
$CheckBoxFolders.location        = New-Object System.Drawing.Point(10,55)
$CheckBoxFolders.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$GroupBoxMail                    = New-Object system.Windows.Forms.Groupbox
$GroupBoxMail.height             = 56
$GroupBoxMail.width              = 586
$GroupBoxMail.text               = "Отправить по почте"
$GroupBoxMail.location           = New-Object System.Drawing.Point(18,159)

$CheckBoxMailOn                  = New-Object system.Windows.Forms.CheckBox
$CheckBoxMailOn.text             = "Ага"
$CheckBoxMailOn.AutoSize         = $false
$CheckBoxMailOn.width            = 50
$CheckBoxMailOn.height           = 20
$CheckBoxMailOn.location         = New-Object System.Drawing.Point(15,25)
$CheckBoxMailOn.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LabelMailAddress                = New-Object system.Windows.Forms.Label
$LabelMailAddress.text           = "Почтовый ящик"
$LabelMailAddress.AutoSize       = $true
$LabelMailAddress.width          = 25
$LabelMailAddress.height         = 10
$LabelMailAddress.location       = New-Object System.Drawing.Point(74,25)
$LabelMailAddress.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBoxMail                     = New-Object system.Windows.Forms.TextBox
$TextBoxMail.multiline           = $false
$TextBoxMail.text                = "media@inplay.pro"
$TextBoxMail.width               = 385
$TextBoxMail.height              = 20
$TextBoxMail.location            = New-Object System.Drawing.Point(187,22)
$TextBoxMail.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ButtonSchedulePath              = New-Object system.Windows.Forms.Button
$ButtonSchedulePath.text         = "Расписание"
$ButtonSchedulePath.width        = 125
$ButtonSchedulePath.height       = 30
$ButtonSchedulePath.location     = New-Object System.Drawing.Point(18,10)
$ButtonSchedulePath.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ButtonAction              = New-Object system.Windows.Forms.Button
$ButtonAction.text         = "Собрать падлу"
$ButtonAction.width        = 125
$ButtonAction.height       = 30
$ButtonAction.location     = New-Object System.Drawing.Point(480,220)
$ButtonAction.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$main.controls.AddRange(@($TextBox1,$GroupboxPrefixes,$GroupboxOptions,$GroupBoxMail,$ButtonSchedulePath,$ButtonAction))
$GroupboxPrefixes.controls.AddRange(@($ReklamaCheckBox,$JingleCheckBox,$BilboardCheckBox,$CustomCheckBox,$TextBox2))
$GroupboxOptions.controls.AddRange(@($CheckBoxPlaylist,$CheckBoxFolders))
$GroupBoxMail.controls.AddRange(@($CheckBoxMailOn,$LabelMailAddress,$TextBoxMail))

# global integers

$PredefinedCheckboxes = $ReklamaCheckBox, $JingleCheckBox, $BilboardCheckBox 
$filePath 

$ButtonSchedulePath.Add_Click({
    $filePath = Get-FilePath
    Set-TextBoxText $TextBox1 $filePath
})

#main function

$ButtonAction.Add_Click({

    #get content of schedule file

    $filePath = $TextBox1.Text
    
    if($filePath -ne $null){
      $schedule = Get-IniContent($filePath)   
    }else{
      Write-Host "Choose schedule file!"  
    }

    #get patterns which we want to reveal in schedule

    $Pattern = Get-PredefinedPatterns $PredefinedCheckboxes

    if($CustomCheckBox.Checked){
        $CustomPattern = Get-UserPattern
    }


    #parse schedule file

    forEach($key in $schedule.Keys){
        
        $path = $schedule[$key].FileName

        if($Pattern -ne "()" -and $path -match $Pattern){
            Write-Host $path
        }

        if($CustomPattern -ne $null -and $path -match $CustomPattern){
            Write-Host $path
        }
    }

    

})


$main.ShowDialog()



function Get-FilePath(){
    $openFileDialog = New-Object windows.forms.openfiledialog

    $openFileDialog.initialDirectory = $env:APPDATA
    $openFileDialog.title = "Выберите файл расписания"
    $openFileDialog.filter = "Schedule file (*.sdl)| *.sdl"
    
    $result = $openFileDialog.ShowDialog()  
    
    if($result -eq "OK") {

        return $openFileDialog.FileName  

    }else{

        Write-Host "Cancelled"

    }  
}



function Set-TextBoxText($textbox, $value){

    $textbox.Text = $value
}

function Get-UserPattern(){

    $customPattern = $TextBox2.text
    return $customPattern 
}

function Get-PredefinedPatterns($checkboxes){


    $Patterns = @()

    forEach($checkbox in $checkboxes){

        if($checkbox.Checked){

            $Patterns += $checkbox.text.split("/")[0]
            $Patterns += $checkbox.text.split("/")[1]
        }
    }

    $CombinedPatterns = '(' + ($Patterns -join ')|(' ) + ')'

    return $CombinedPatterns
}


function Get-IniContent($filePath){

$ini=@{}

    switch -Regex -File $filePath
    {
      “^\[(.+)\]” # Section - начало строки [ должен быть минимум один символ, конец ]
      {
        $section = $matches[1] # сохраняем то, что находится между [] ак хэш-таблицу
        $ini[$section] = @{}
        $CommentCount = 0
      }
      “^(;.*)$” # Comment
      {
        $value = $matches[1]
        $CommentCount = $CommentCount + 1
        $name = “Comment” + $CommentCount
        $ini[$section][$name] = $value
      }
      “(.+?)\s*=(.*)” # Key 
      {
        $name,$value = $matches[1..2]
        $ini[$section][$name] = $value
      }
    }

    return $ini
}

function Format-Name($string){
    $string.Substring($string.IndexOf("-")+2)
}

