#requires -version 2
<#
.SYNOPSIS
  This script generates a CSV file with asset information. 
.DESCRIPTION
  The script will create a CSV file about the local computer or a remote computer if its hostname is provided as argument (see below)
.PARAMETER <hostname>
  If provided, generates the report about a remote computer.
.INPUTS
  None
.OUTPUTS
  Comma-separated file <hostname>.csv 
.NOTES
  Version:        1.0
  Author:         Sergio Acosta
  Creation Date:  17/01/2016
  Purpose/Change: -
.EXAMPLE
  csvInventory.ps1 localhost
#>


