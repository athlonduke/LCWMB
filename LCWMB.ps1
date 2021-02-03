[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

# CHANGE THESE TO YOUR STUFF
$From = "YOUREMAILADDRESS"
$To = "YOURSERVICECONNECTORADDRESS"
$SMTPServer = "YOUREMAILSERVER"
$SMTPPort = "YOUREMAILSERVERPORT"


function MakeTextBox {
    param ($name, $xcoord, $ycoord, $width) 
    $objtextbox = New-Object System.Windows.Forms.textbox
    $objtextbox.left = $xcoord 
    $objtextbox.top = $ycoord 
    $objtextbox.name = $name
    $objtextbox.Width = $width
    $objForm.Controls.Add($objtextbox)
}
function MakeLabel {
    param ($text, $xcoord, $ycoord) 
    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.left = $xcoord 
    $objLabel.top = $ycoord 
    $objLabel.Text = $text
    $objlabel.width = 150
    $objForm.Controls.Add($objLabel)
}

function MakeButton {
    param ($name, $text, $xcoord, $ycoord) 
    $name = New-Object System.Windows.Forms.Button
    $name.Left = $xcoord
    $name.top = $ycoord
    $name.width = 90
    $name.Text = $text
    $name.Add_Click( 
        {
            $objform.controls["ticketNotes"].text = $($this).Text
            
        }
    )
    $objForm.Controls.Add($name)         
}

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Lazy CWM Bastard"
$objForm.Size = New-Object System.Drawing.Size(400,240) 
$objForm.name = "form"
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$x=$objListBox.SelectedItem;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})


MakeLabel -text "Ticket #" -xcoord 10 -ycoord 20
MakeTextBox -name "ticketNumber" -xcoord 10 -ycoord 45 -width 60

MakeLabel -text "Hours" -xcoord 160 -ycoord 20
MakeTextBox -name "ticketTime" -xcoord 160 -ycoord 45 -width 40

MakeLabel -text "Ticket Notes" -xcoord 10 -ycoord 80
MakeTextBox -name "ticketNotes" -xcoord 10 -ycoord 105 -width 200

# THIS SECTION IS WHERE YOU CAN MAKE YOUR OWN BUTTONS
MakeButton -name "btnAdmin" -text "Admin time" -xcoord 10 -ycoord 130
MakeButton -name "btnTech" -text "Tech consult" -xcoord 10 -ycoord 155
MakeButton -name "btnClose" -text "!!status:>Closed!!" -xcoord 100 -ycoord 130

# put in default of 0.25 hours
$objform.controls["ticketTime"].text = .25

$StaticButton = New-Object System.Windows.Forms.Button
$StaticButton.Left = 250
$StaticButton.top = 160
$StaticButton.width = 90
$StaticButton.Text = "Send Update"
$StaticButton.Add_Click(
    {
        $finalTicketNumber = ($objform.controls["ticketNumber"].text).TrimEnd()
        $finalTicketTime = ($objform.controls["ticketTime"].text)  + "!!Time:" + ($objform.controls["ticketTime"].text) + "!!"
        $finalTicketNotes = ($objform.controls["ticketNotes"].text).tostring()
        Write-Host $finalTicketNumber $finalTicketTime $finalTicketNotes

        $Subject = "Ticket#" + $finalTicketNumber + "/ Update"
        $Body  = "Ticket# - " + $finalTicketNumber + "<br>"
        $Body += "Time - " + $finalTicketTime + "<br>"
        $Body += "Notes - " + $finalTicketNotes + "<br>"

        try {
            $emailResult = Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort -Verbose
        } catch [System.Exception] {
            "Failed to send email: {0}" -f  $_.Exception.Message
        }
        MakeLabel -text "Ticket $finalTicketNumber Updated" -xcoord 250 -ycoord 90
    }
    
)
$objForm.Controls.Add($StaticButton)  

$objForm.Topmost = $True
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

