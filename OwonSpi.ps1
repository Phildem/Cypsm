

<#
Cypsm

This powerShell script is just a quick example to demonstrate that it is very easy to pilot SPI devices with PowerShell
You can use it as it is but the idea is of course to adapt it to your needs.

It has been developped and tested to drive a Owon Programable Power supply SPM6103

The manual can be downloaded here : https://files.owon.com.cn/probook/SPM_Series_User_Manual.pdf
The programing manual here : https://files.owon.com.cn/software/Application/SPM_Series_programming_manual.pdf

The USB interface create a port COM (See on your PC wich one and set the $portName variable.)

To be simple I havn't add any error handling.


MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.


THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

#>

#****************************************************************************
# Funstion definition
#****************************************************************************

#Open the serial Port _________________________________________________________________________
function OpenPort {

    $serialPort.PortName = $portName
    $serialPort.BaudRate = 115200
    $serialPort.Parity = [System.IO.Ports.Parity]::None
    $serialPort.DataBits = 8
    $serialPort.StopBits = [System.IO.Ports.StopBits]::One


    $serialPort.Open()
}

#Close the serial Port _________________________________________________________________________
function ClosePort {
    $serialPort.Close()
}

#_____________________________________________________________________________________________
function DisplayData {
    OpenPort
    $serialPort.WriteLine("*IDN?")
    $receivedData = $serialPort.ReadLine()
    Write-Host "$receivedData"
    ClosePort
}

#  Allow to set PSu Voltage, Current limit, OVP and OCP ________________________________________
function PsuSet {

    param (
        [string]$VOut="*",
        [string]$ILim="*",
        [string]$OVP="*",
        [string]$OCP="*"
    )

    OpenPort

    if ($OVP -ne '*') {$serialPort.WriteLine("VOLT:LIM $OVP")}
    if ($OCP -ne '*') {$serialPort.WriteLine("CURR:LIM $OCP")}
    if ($VOut -ne '*') {$serialPort.WriteLine("VOLT $VOut")}
    if ($ILim -ne '*') {$serialPort.WriteLine("CURR $ILim")}
    ClosePort
}

#Allow to switch psu output On or Off (cmd is ON or OFF ) ___________________________________________________________
function PsuOut {

    param (
        [string]$Cmd      
    )

    OpenPort
    $serialPort.WriteLine("OUTP $Cmd")
    ClosePort
}


#_ test for key s pressed, work ISE and Console ________________________________________________
function TestStopKey(){

 [bool]$KeyPress=([PsOneApi.Keyboard]::GetAsyncKeyState($key) -eq -32767)
 if ($KeyPress){
    Write-Host "n`**** Process interupted ****" 
 }

return $KeyPress
}


# Generate an voltage ramp _______________________________________________________________________
function PsuVar {

    param (
       [decimal]$VMin,
       [decimal]$VMax,
       [decimal]$Inc=0.1,
       [decimal]$Delay=500
    )


     Write-Host "Ramp from $VMin to $VMax Step $Inc every $Delay" 
     Write-Host "Press 's' to stop" 


    OpenPort

    $serialPort.WriteLine("VOLT:LIM $VMax")

    for ($VOut = $VMin; $VOut -le $VMax; $VOut += $Inc) {

        if (TestStopKey) {
        break
        }

        $serialPort.WriteLine("VOLT $VOut")
        $serialPort.WriteLine("VOLT?")
        Write-Host $serialPort.ReadLine()
        Start-Sleep -Milliseconds $Delay
    }

    ClosePort
}


# Generate random voltage variations ______________________________________________________
function PsuRnd {

    param (
       [decimal]$Var=0,
       [decimal]$Delay=0
    )


     OpenPort

     $serialPort.WriteLine("VOLT?")
     [decimal]$VNom=$serialPort.ReadLine()


     if ($Var -eq 0){

        Write-Host "Current output voltage is $VNom V " 

         $Var = Read-Host "`nMax variation in Volt "

     }

      if ($Var -eq 0){
        return
      }


    if ($Delay -eq 0){

         $Delay = Read-Host "`nDelay between changes in ms"
     }


     [decimal]$VMin=$VNom-$Var
     [decimal]$VMax=$VNom+$Var
     if ($VMin -lt 0){
        $VMin=0;
     }   

     if ($VMax -gt 60){
        $VMax=60;
     }   


     Write-Host "Will randomly set VOut betwwen $VMin et $VMax every $Delay ms " 
     Write-Host "Press 's' to stop" 


    $serialPort.WriteLine("VOLT:LIM $VMax")

    do {

        if (TestStopKey) {
        break
        }

        $VOut = Get-Random -Minimum $VMin -Maximum $VMax
        $VOut = [math]::Round($VOut, 2)

        $serialPort.WriteLine("VOLT $VOut")
        $serialPort.WriteLine("VOLT?")
        Write-Host $serialPort.ReadLine()
        Start-Sleep -Milliseconds $Delay
    } while(1)

    ClosePort
}





#_____________________________________________________________________________________________
function DisplayMenu {
    Clear-Host
    Write-Host "Cypsm 1.0 OWON SPM Série SPI control by Cyrob "
    Write-Host "****************************************`n"
    Write-Host "1. Get device Info"
    Write-Host "2. Set for 5V 1A OVP 5.5 OCP 1.1"
    Write-Host "3. Set for 12V 2A OVP 15 OCP 2.1"
    Write-Host "R. Ramp 0 -> 60V by 0.5V increment "
    Write-Host "V. Simulate variation in current voltage "
    Write-Host ""
    Write-Host "O. Output On"
    Write-Host "F. Output Off"
    Write-Host ""
    Write-Host "Q. Quit"
}


#****************************************************************************
# Main program start here
#****************************************************************************

# Init abort key test
$key = [Byte][Char]'S'    
$Signature = @'
    [DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] 
    public static extern short GetAsyncKeyState(int virtualKeyCode); 
'@
Add-Type -MemberDefinition $Signature -Name Keyboard -Namespace PsOneApi
    

# Instantiate Serial port
$serialPort = New-Object System.IO.Ports.SerialPort

# SET Your serial port name HERE
$portName = "COM9"


do {
    DisplayMenu
    $choix = Read-Host "`nChoice "

    $Needkey=0  #Set to 1 if the command display result

    Write-Host ""

    switch ($choix) {
        1 {
            DisplayData
            $Needkey=1
        }

        2 {
            PsuSet -VOut 5 -ILim 1 -OVP "5.5" -OCP "1.1"
        }

        3 {
            PsuSet -VOut 12 -ILim 2 -OVP "15" -OCP "2.1"
        }


        'r' {
            PsuVar -VMin 0 -VMax 60 -Inc 0.5
        }

        'v' {
           PsuRnd
        }

        'o' {
           PsuOut -Cmd "ON"
        }

        'f' {
           PsuOut -Cmd "OFF"
        }


        'q' {
            Write-Host "Good bye..."
        }
        default {
            Write-Host "Bad choice, try again !"
            $Needkey=1
        }
    }
    if ($Needkey -ne '0') {
        Read-Host "`n`nHit any key to continue..."
    }
} while ($choix -ne 'q')


