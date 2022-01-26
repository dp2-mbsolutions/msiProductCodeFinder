
###############################################################################################################################
#                                                             Visual Elements                                                 #
###############################################################################################################################

[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
$ProductCodeFinder = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.OpenFileDialog]$openFileDialog1 = $null
[System.Windows.Forms.Button]$button2 = $null
[System.Windows.Forms.TextBox]$textBox1 = $null
[System.Windows.Forms.TextBox]$textBox2 = $null
[System.Windows.Forms.Button]$button3 = $null
[System.Windows.Forms.Label]$label1 = $null
[System.Windows.Forms.Button]$button1 = $null
function InitializeComponent
{
$openFileDialog1 = (New-Object -TypeName System.Windows.Forms.OpenFileDialog)
$button2 = (New-Object -TypeName System.Windows.Forms.Button)
$textBox1 = (New-Object -TypeName System.Windows.Forms.TextBox)
$textBox2 = (New-Object -TypeName System.Windows.Forms.TextBox)
$button3 = (New-Object -TypeName System.Windows.Forms.Button)
$label1 = (New-Object -TypeName System.Windows.Forms.Label)
$ProductCodeFinder.SuspendLayout()
#
#openFileDialog1
#
$openFileDialog1.Filter = [System.String]'msi files (*.msi)|*.msi|All files (*.*)|*.*'
$openFileDialog1.InitialDirectory = [System.String]'C:\'
$openFileDialog1.Title = [System.String]'Browse MSI Files'
#
#button2
#
$button2.BackColor = [System.Drawing.Color]::AliceBlue
$button2.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$button2.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]428,[System.Int32]34))
$button2.Name = [System.String]'button2'
$button2.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]106,[System.Int32]21))
$button2.TabIndex = [System.Int32]0
$button2.Text = [System.String]'Browse'
$button2.UseVisualStyleBackColor = $false
$button2.add_Click{(Select-MSI)}
#
#textBox1
#
$textBox1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]142,[System.Int32]35))
$textBox1.Name = [System.String]'textBox1'
$textBox1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]280,[System.Int32]20))
$textBox1.TabIndex = [System.Int32]1
#
#textBox2
#
$textBox2.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]27,[System.Int32]148))
$textBox2.Multiline = $true
$textBox2.AcceptsReturn = $true
$textBox2.Name = [System.String]'textBox2'
$textBox2.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]526,[System.Int32]94))
$textBox2.TabIndex = [System.Int32]2
#
#button3
#
$button3.BackColor = [System.Drawing.Color]::AliceBlue
$button3.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]142,[System.Int32]71))
$button3.Name = [System.String]'button3'
$button3.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]280,[System.Int32]62))
$button3.TabIndex = [System.Int32]3
$button3.Text = [System.String]'Display MSI Info'
$button3.UseVisualStyleBackColor = $false
$button3.add_Click{(Display-Info)}
#
#label1
#
$label1.AutoSize = $true
$label1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Microsoft Sans Serif',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$label1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]24,[System.Int32]42))
$label1.Name = [System.String]'label1'
$label1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]112,[System.Int32]13))
$label1.TabIndex = [System.Int32]4
$label1.Text = [System.String]'Select the .msi file'
#
#ProductCodeFinder
#
$ProductCodeFinder.BackColor = [System.Drawing.SystemColors]::ActiveCaption
$ProductCodeFinder.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]577,[System.Int32]252))
$ProductCodeFinder.Controls.Add($label1)
$ProductCodeFinder.Controls.Add($button3)
$ProductCodeFinder.Controls.Add($textBox2)
$ProductCodeFinder.Controls.Add($textBox1)
$ProductCodeFinder.Controls.Add($button2)
$ProductCodeFinder.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$ProductCodeFinder.Name = [System.String]'ProductCodeFinder'
$ProductCodeFinder.Text = [System.String]'.msi ProductCodeFinder'
$ProductCodeFinder.ResumeLayout($false)
$ProductCodeFinder.PerformLayout()
Add-Member -InputObject $ProductCodeFinder -Name base -Value $base -MemberType NoteProperty
Add-Member -InputObject $ProductCodeFinder -Name openFileDialog1 -Value $openFileDialog1 -MemberType NoteProperty
Add-Member -InputObject $ProductCodeFinder -Name button2 -Value $button2 -MemberType NoteProperty
Add-Member -InputObject $ProductCodeFinder -Name textBox1 -Value $textBox1 -MemberType NoteProperty
Add-Member -InputObject $ProductCodeFinder -Name textBox2 -Value $textBox2 -MemberType NoteProperty
Add-Member -InputObject $ProductCodeFinder -Name button3 -Value $button3 -MemberType NoteProperty
Add-Member -InputObject $ProductCodeFinder -Name label1 -Value $label1 -MemberType NoteProperty
Add-Member -InputObject $ProductCodeFinder -Name button1 -Value $button1 -MemberType NoteProperty
}
. InitializeComponent




###############################################################################################################################
#                                                             Function Elements                                               #
###############################################################################################################################

#Hide Console 

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

#Get MSI Information Function

<#

.LINK
   http://blog.kmsigma.com/

.LINK
   https://msdn.microsoft.com/en-us/library/windows/desktop/aa370905(v=vs.85).aspx

.LINK
   http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/

.NOTES
   Heavily Infuenced by http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/
   I did not write this function, it is a fucntion from the above links

#>

function Get-MsiInformation
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$false,
                   ConfirmImpact='Medium')]
    [Alias("gmsi")]
    Param(
        [parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage = "Provide the path to an MSI")]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]]$Path,
 
        [parameter(Mandatory=$false)]
        [ValidateSet( "ProductCode", "Manufacturer", "ProductName", "ProductVersion", "ProductLanguage" )]
        [string[]]$Property = ( "ProductCode", "Manufacturer", "ProductName", "ProductVersion", "ProductLanguage" )
    )

    Begin
    {
        # Do nothing for prep
    }
    Process
    {
        
        ForEach ( $P in $Path )
        {
            if ($pscmdlet.ShouldProcess($P, "Get MSI Properties"))
            {            
                try
                {
                    Write-Verbose -Message "Resolving file information for $P"
                    $MsiFile = Get-Item -Path $P
                    Write-Verbose -Message "Executing on $P"
                    
                    # Read property from MSI database
                    $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
                    $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MsiFile.FullName, 0))
                    
                    # Build hashtable for retruned objects properties
                    $PSObjectPropHash = [ordered]@{File = $MsiFile.FullName}
                    ForEach ( $Prop in $Property )
                    {
                        Write-Verbose -Message "Enumerating Property: $Prop"
                        $Query = "SELECT Value FROM Property WHERE Property = '$( $Prop )'"
                        $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
                        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
                        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
                        $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
 
                        # Return the value to the Property Hash
                        $PSObjectPropHash.Add($Prop, $Value)

                    }
                    
                    # Build the Object to Return
                    $Object = @( New-Object -TypeName PSObject -Property $PSObjectPropHash )
                    
                    # Commit database and close view
                    $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
                    $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)           
                    $MSIDatabase = $null
                    $View = $null
                }
                catch
                {
                    Write-Error -Message $_.Exception.Message
                }
                finally
                {
                    Write-Output -InputObject @( $Object )
                }
            } # End of ShouldProcess If
        } # End For $P in $Path Loop

    }
    End
    {
        # Run garbage collection and release ComObject
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        [System.GC]::Collect()
    }
}
<#
End of Function
#>



#Button Functions

function Select-MSI 
{
 $Textbox1.text = ""
 $openFileDialog1.ShowDialog()

 if ($openFileDialog1.FileName.Length -gt '0') {
    
    $Textbox1.Text = $openFileDialog1.FileName

    }
    
}

function Display-Info 
{
 
 $textbox2.text = @()
 $productname = (Get-MsiInformation -Path $textbox1.Text -Property ProductName).productname
 $productcode = (Get-MsiInformation -Path $textbox1.Text -Property ProductCode).productcode
 $manufacturer = (Get-MsiInformation -Path $textbox1.Text -Property Manufacturer).Manufacturer
 $productVersion = (Get-MsiInformation -Path $textbox1.Text -Property ProductVersion).ProductVersion
 $productLanguage = (Get-MsiInformation -Path $textbox1.Text -Property ProductLanguage).ProductLanguage 
 $textbox2.text = "Product Name: $productname`r`nProductCode: $productcode`r`nManufacturer: $manufacturer`r`nProductVersion: $productVersion`r`nProductLanguage: $productLanguage"
    
}







###############################################################################################################################
#                                                             Start Form                                                      #
###############################################################################################################################


$ProductCodeFinder.ShowDialog()