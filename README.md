Powershell Printer Forms Module 
===========

Powershell module to administer windows printer forms.

## Example Usage

Get all of the forms installed on the computer:

	Windows PowerShell
	Copyright (C) 2016 Microsoft Corporation. All rights reserved.

	PS C:\Drop> Import-Module .\PowershellPrinterFormsModule.dll
	PS C:\Drop> Get-SystemForm | ft

	IsUserForm Name                           PageSize        ImageableArea       Margin  MarginMM MarginInches PageSizeMM        PageSizeInches
	---------- ----                           --------        -------------       ------  -------- ------------ ----------        --------------
		 False Letter                         215900,279400   0,0,215900,279400   0,0,0,0 0,0,0,0  0,0,0,0      215.9,279.4       8.5,11
		 False Letter Small                   215900,279400   0,0,215900,279400   0,0,0,0 0,0,0,0  0,0,0,0      215.9,279.4       8.5,11
		 False Tabloid                        279400,431800   0,0,279400,431800   0,0,0,0 0,0,0,0  0,0,0,0      279.4,431.8       11,17
		 False Ledger                         431800,279400   0,0,431800,279400   0,0,0,0 0,0,0,0  0,0,0,0      431.8,279.4       17,11
		 False Legal                          215900,355600   0,0,215900,355600   0,0,0,0 0,0,0,0  0,0,0,0      215.9,355.6       8.5,14
    *** Snipped for brevity ***

	PS C:\Drop>

Get a specifc form:

	PS C:\Drop> Get-SystemForm 'Demo User Form'
	
	Name                : Demo User Form
	IsUserForm          : True
	PageSize            : 101600,127000
	PageSizeMM          : 101.6,127
	PageSizeInches      : 4,5
	ImageableArea       : 6350,12700,88900,101600
	ImageableAreaMM     : 6.35,12.7,88.9,101.6
	ImageableAreaInches : 0.25,0.5,3.5,4
	Margin              : 6350,12700,6350,12700
	MarginMM            : 6.35,12.7,6.35000000000001,12.7
	MarginInches        : 0.25,0.5,0.25,0.5

	PS C:\Drop>

Add new forms:

	PS C:\Drop> Add-SystemForm -Name 'Demo User Form try 1' -Units Inches -Size '4,5'
	PS C:\Drop> Add-SystemForm -Name 'Demo User Form try 2' -Units Inches -Size '4,5' -Margin '0.25,0.5'
	PS C:\Drop> Add-SystemForm -Name 'Demo User Form try 3' -Units Millimeters -Size '80,50' -Margin '10,10,0,0'

Delete all user forms:

	PS C:\Drop> Get-SystemForm | ?{ $_.IsUserForm } | Remove-SystemForm

