
# automssqlbackup.ps1

#=====================================================================
# MSSQL Backup Script
# Version 0.32
# http://www.devio.at/index.php/automssqlbackup
# (c) 2009-2012 devio IT Services
# 2020 updated SQL Server versions
# support@devio.at
 
#=====================================================================
# This program is inspired by and based on automysqlbackup.sh:
#
# MySQL Backup Script
# Ver. 2.5 - http://sourceforge.net/projects/automysqlbackup/
# Copyright (c) 2002-2003 wipe_out@lycos.co.uk
# and therefore subject to the GNU GPL:

# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 2 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

#=====================================================================
# Set the following variables to your system needs
# (Detailed instructions around variables)
#=====================================================================


# to execute this .ps1 script from command line or batch file, run
# Set-ExecutionPolicy RemoteSigned 
# in powershell (admin mode) first to allow script execution


# set $mspath to localized and SQL Server version-specific program files path
# requires trailing backslash

# SQL Server 2005

# en: $mspath = "C:\Program Files\Microsoft SQL Server\90\SDK\Assemblies\"
# de: $mspath = "C:\Programme\Microsoft SQL Server\90\SDK\Assemblies\"

# SQL Server 2008
# en: $mspath = "C:\Program Files\Microsoft SQL Server\100\SDK\Assemblies\"
# de: $mspath = "C:\Programme\Microsoft SQL Server\100\SDK\Assemblies\"

# SQL Server 2012
# en: $mspath = "C:\Program Files\Microsoft SQL Server\110\SDK\Assemblies\"

# SQL Server 2014
# en win64: $mspath = "C:\Program Files (x86)\Microsoft SQL Server\120\SDK\Assemblies\"

# SQL Server 2016
# en: $mspath = "C:\Program Files\Microsoft SQL Server\130\SDK\Assemblies\"
# en: $mspath = "C:\Program Files\Microsoft SQL Server\130\Tools\Binn\ManagementStudio\"
# en win64: $mspath = "C:\Program Files (x86)\Microsoft SQL Server\130\Tools\Binn\ManagementStudio\"

# SSMS 2017
# en win64: $mspath = "C:\Program Files (x86)\Microsoft SQL Server\140\Tools\Binn\ManagementStudio\"

# SSMS 18
# en win64: $mspath = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\"

&{
# load SMO assemblies

	$dummy = [System.Reflection.Assembly]::LoadFrom($mspath + "Microsoft.SqlServer.ConnectionInfo.dll")
	$dummy = [System.Reflection.Assembly]::LoadFrom($mspath + "Microsoft.SqlServer.Smo.dll")
	
# additional SMO assembly required for 2008 and 2012 (not available in 2005!)

	$dummy = [System.Reflection.Assembly]::LoadFrom($mspath + "Microsoft.SqlServer.SmoExtended.dll")

# load SharpZipLib

	$dummy = [System.Reflection.Assembly]::LoadFrom("C:\path-to\ICSharpCode.SharpZipLib.dll")

	trap {
		$_;
		break
	}
}

# begin of config section

# login info. set to "" for windows authorization, SQL authorization otherwise
$username = ""	# "username"
$password = ""	# "password"

# db server (including \instance name), db names
# set $dbnames = "" for all databases
# set $dbnames = "- db1 db2" (leading minus sign) for all databases *except* listed
$dbhost = "localhost"
$dbnames = ""		# e.g. "db1 db2 db3"

# backup dir (for sql server)
$backupdir = "C:\my-backup-dir"

$remotebackupdir = $backupdir
$logdir = $backupdir

# if this script runs on the same server as the database, the $...dir variables
# have the same value
# if this script performs backups on a remote server, $backupdir and
# $remotebackupdir have different values, but have to point to the same
# physical directory. $backupdir contains the path as valid from SQL server,
# $remotebackupdir contains the remote path as visible from the machine the
# script is running on, typically in UNC notation

# if remote backup, set remote backup dir (UNC), else leave same
#$remotebackupdir = "\\host\share\backupdir"

# if log dir is not backup root dir, set here
#$logdir = "C:\my-log-dir"

# mail settings
$emailFrom = "backup@example.com"
$emailTo = "me@example.com"
$subject = "automssqlbackup log"
$smtpServer = ""	# set to activate email "smtp.example.com"


# day of week for weekly backup, monday = 1
$doweekly = 6

# daily full backup (default = 0 = incremental)
$dailyfull = 0

# each database in separate directory
$sepdir = 1

# compress to .zip folder (xp and higher)
$compress = 1

# keep latest backup in "latest" directory
$latest = 1

# end of config section

#=====================================================================
# Scheduling
#
# - Powershell must be allowed to execute scripts, this is done by
#	Set-ExecutionPolicy RemoteSigned
# - place .ps1 and .cmd file in a directory
# - modify .cmd to point to the .ps1
# - adjust .ps1 to your environment
# - point Explorer to Scheduled Task, Add Scheduled Task,
#   select the .cmd, choose daily execution, select time
# - selected login credentials must be able to execute .cmd and .ps1
#	and connect to SQL Server if you omit username/password
# - after creating the Task, right-click and select Run to make sure
#   it is setup correctly
# - check log files
#

# taken from automysqlbackup.sh:
#
#=====================================================================
# Backup Rotation..
#=====================================================================
#
# Daily Backups are rotated weekly..
# Weekly Backups are run by default on Saturday..
# Weekly Backups are rotated on a 5 week cycle..
# Monthly Backups are run on the 1st of the month..
# Monthly Backups are NOT rotated automatically...
# It may be a good idea to copy Monthly backups offline or to another
# server..
#
#=====================================================================
# Please Note!!
#=====================================================================
#
# I take no responsibility for any data loss or corruption when using
# this script..
# This script will not help in the event of a hard drive crash. If a 
# copy of the backup has not be stored offline or on another PC..
# You should copy your backups offline regularly for best protection.
#

#=====================================================================
# Restoring
#=====================================================================
#
# Use SQL Server's Enterprise Manager or Management Studio to 
# restore database.
#

#=====================================================================
# backup files names:
#   separate dir:
# 	daily:		$backupdir\daily\$dbname\$dbname-d-yyyy-mm-dd.bak
# 	weekly:		$backupdir\weekly\$dbname\$dbname-ww-yyyy-mm-dd.bak
# 	monthly:	$backupdir\monthly\$dbname\$dbname-yyyy-mm-dd.bak
#   same dir:
# 	daily:		$backupdir\daily\$dbname-d-yyyy-mm-dd.bak
# 	weekly:		$backupdir\weekly\$dbname-ww-yyyy-mm-dd.bak
# 	monthly:	$backupdir\monthly\$dbname-yyyy-mm-dd.bak
#
#	latest:		$backupdir\latest\$dbname-pattern.ext
#
#   $compress adds ".zip" to filenames

#=====================================================================
# Change Log
# 0.32	200821	add assembly paths for SSMS 2014, 2016, 2017, SSMS 18
# 0.31	120719	load assemblies for SQL SMO 2005, 2008, 2012
# 0.30	090731	log exception message via $error[0]
# 0.29	090407	long filesize, StatementTimeout, exception handler
# 0.28	090223	Table of databases and results as HTML in notification mail
# 0.27	090217	Error indicator and backup schedule in mail subject
# 0.26	090211	SharpZipLib, exclude databases from "all"
# 0.25	090209	initial release

#=====================================================================
# From here on, modify only if you know what you do:
#
#=====================================================================
# function definitions

function CreateDir($dir)
{
	if ($dir -eq "") {
		$dir = $remotebackupdir
	}
	else {
		$dir = [System.IO.Path]::Combine($remotebackupdir, $dir);
	}
	
	if(![System.IO.Directory]::Exists($dir)) { 
		$dummy = [System.IO.Directory]::CreateDirectory($dir);
	}
}

# $srv is global identifying target server

function BackupDB($dbname, $backupfile, $full)
{
	$remotefile = [System.IO.Path]::Combine($remotebackupdir, $backupfile);
	$backupfile = [System.IO.Path]::Combine($backupdir, $backupfile);
	
	if ([System.IO.File]::Exists($remotefile)) {
		[System.IO.File]::Delete($remotefile)
	}

	$script:bkstat = "ok"

	#$db = New-Object Microsoft.SqlServer.Management.Smo.Database
	$db = $srv.Databases[$dbname]
	Write-Host ($dbname + " connections: " + $db.ActiveConnections.ToString())
	
	$bk = New-Object Microsoft.SqlServer.Management.Smo.Backup
	$bk.Action = [Microsoft.SqlServer.Management.Smo.BackupActionType]::Database
	$bk.Database = $dbname
	$bk.Initialize = $True
	$bk.Incremental = -not $full
	
	$bk.Devices.AddDevice($backupfile, [Microsoft.SqlServer.Management.Smo.DeviceType]::File)
	Write-Host ("backing up " + $dbname)
	$bk.SqlBackup($srv)
	Write-Host ("created " + $backupfile)

	$fi = New-Object System.IO.FileInfo($remotefile)
	if ($fi.Length) {
		$dbstat[$dbname] = ($script:bkstat, $backupfile, $remotefile, $fi.Length, $null)
	}
	else {
		$dbstat[$dbname] = ($script:bkstat, $backupfile, $remotefile, "missing", $null)
	}
	
	trap {
		Write-Host "backup failed"
		$_;
		$script:bkstat = $_
		if ($_.InnerException) {
			$_.InnerException;
		}
		Write-Host "exception message"
		Write-Host $error[0].Exception.ToString();
		continue
	}
}

function Compress($dbname, $backupfile)
{
	$filename = [System.IO.Path]::Combine($remotebackupdir, $backupfile)

	if ([System.IO.File]::Exists($filename)) 
	{
		if ($compress) 
		{
			$zipname = [System.IO.Path]::Combine($remotebackupdir, $backupfile) + ".zip"
			if ([System.IO.File]::Exists($zipname)) {
				[System.IO.File]::Delete($zipname)
			}

			#$zip = [ICSharpCode.SharpZipLib.Zip.ZipFile]::Create($zipname)
			#$zip.BeginUpdate()
			#$zip.Add($filename)
			#$zip.CommitUpdate()
			#$zip.Close()
			
			$filenamedir = $filename.Substring(0, `
				$filename.Length - [System.IO.Path]::GetFileName($filename).Length - 1)
				
			$zip = New-Object ICSharpCode.SharpZipLib.Zip.FastZip
			$zip.CreateZip($zipname, $filenamedir, $false, "\.bak$")

			#CreateZip(string zipFileName, string sourceDirectory, 
			#bool recurse, string fileFilter, string directoryFilter)

			[System.IO.File]::Delete($filename)

			$backupfile = $backupfile + ".zip"	
			
			$fi = New-Object System.IO.FileInfo($zipname)
			if ($fi.Length) {
				$dbs = $dbstat[$dbname]
				$dbs[2] = $zipname
				$dbs[4] = $fi.Length
				$dbstat[$dbname] = $dbs
			}
		}

		if ($latest) {
			$dummy = Copy-Item ([System.IO.Path]::Combine($remotebackupdir, $backupfile)) `
				([System.IO.Path]::Combine($remotebackupdir, "latest"))
		}		
	}

	$backupfile
}

# end of function definitions

$ver = "0.32"

$now = Get-Date

# calculate file names

$ci = [System.Globalization.CultureInfo]::CreateSpecificCulture("no")
$cal = $ci.Calendar
$week = $cal.GetWeekOfYear($now, $ci.DateTimeFormat.CalendarWeekRule, $ci.DateTimeFormat.FirstDayOfWeek)
$remweek = $cal.GetWeekOfYear($now.AddDays(-35), $ci.DateTimeFormat.CalendarWeekRule, $ci.DateTimeFormat.FirstDayOfWeek)

$escdbhost = $dbhost -replace "\\", "_"		# handle server\instance

$logfile = [System.IO.Path]::Combine($logdir, $escdbhost + "-" + $now.ToString("yyyy-MM-dd") + ".log")

# create backup directories

Start-Transcript -path $logfile -force

CreateDir ""
CreateDir "daily"
CreateDir "weekly"
CreateDir "monthly"

if ($latest) {
	CreateDir "latest"
	$b = [System.IO.Path]::Combine($remotebackupdir, "latest")
	Remove-Item ($b + "\*") -recurse -force 
}

# create smo server connection

if ($username -eq "") {
	$conn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection($dbhost)
}
else {
	$conn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection($dbhost, $username, $password)
}
$conn.StatementTimeout = 0

$srv = New-Object Microsoft.SqlServer.Management.Smo.Server($conn)

if (($dbnames -eq "") -or $dbnames.StartsWith("-"))
{
	$dbexclude = ""
	if ($dbnames.StartsWith("-")) {
		$dbexclude = " " + $dbnames.Substring(1) + " "
	}
	
	$dbnames = ""
	
	foreach($db in $srv.Databases | 
		Where-Object { $_.IsAccessible -and -not $_.IsSystemObject } | 
		Where-Object { $dbexclude.IndexOf(" " + $_.Name + " ", [StringComparison]::CurrentCultureIgnoreCase) -eq -1 } |
		Sort-Object { $_.Name })
	{
		if ($dbnames -eq "") {
			$dbnames = $db.Name
		}
		else {
			$dbnames = $dbnames + " " + $db.Name
		}
	}
}

Write-Host "======================================================================"
Write-Host ("AutoMSSQLBackup " + $ver)
#Write-Host "http://www.devio.at/index.php/automssqlbackup"
Write-Host ""
Write-Host ("Backup of Database Server " + $dbhost)
Write-Host ("Databases: " + $dbnames)
Write-Host "======================================================================"
Write-Host ("Backup Start: " + [DateTime]::Now.ToString())

$backupfiles = ""

$dbstat = @{}
$dbstatmonthly = $null

if ($now.Day -eq 1)	# 1st of month => monthly backup
{
	$subject = $subject + " monthly"
	foreach($db in $dbnames.Split(' '))
	{
		Write-Host ("Monthly Backup of " + $db)

		$b = "monthly"	
		if ($sepdir) {
			$b = [System.IO.Path]::Combine($b, $db);
		}
		CreateDir $b
		$b = [System.IO.Path]::Combine($b,
			$db + "-" + $now.ToString("yyyy-MM-dd") + ".bak")

		BackupDB $db $b $true
		$b = Compress $db $b
		$backupfiles = $backupfiles + " " + $b
	}

	$dbstatmonthly = $dbstat
	$dbstat = @{}
}

if ($doweekly -eq $now.DayOfWeek)	# weekly backup
{
	$subject = $subject + " weekly"
}
else
{
	$subject = $subject + " daily"
}

foreach($db in $dbnames.Split(' '))
{
	$b = "daily"	
	$b = [System.IO.Path]::Combine($b, $db);
	CreateDir $b

	if ($doweekly -eq $now.DayOfWeek)	# weekly backup
	{
		Write-Host ("Weekly Backup of " + $db)

		$b = "weekly"	
		if ($sepdir) {
			$b = [System.IO.Path]::Combine($b, $db);
		}
		CreateDir $b

		$remb = [System.IO.Path]::Combine($remotebackupdir, $b)
		Remove-Item ($remb + "\" + $db + "-" + $remweek.ToString("00") + "*") -recurse -force 

		$b = [System.IO.Path]::Combine($b,
			$db + "-" + $week.ToString("00") + "-" + $now.ToString("yyyy-MM-dd") + ".bak")

		BackupDB $db $b $true
		$b = Compress $db $b
		$backupfiles = $backupfiles + " " + $b
	}
	else					# daily backup
	{
		Write-Host ("Daily Backup of " + $db)

		$b = "daily"	
		if ($sepdir) {
			$b = [System.IO.Path]::Combine($b, $db);
		}
		CreateDir $b

		$remb = [System.IO.Path]::Combine($remotebackupdir, $b)
		Remove-Item ($remb + "\" + $db + "-" + ([int]$now.DayOfWeek).ToString(0) + "*") -recurse -force 

		$b = [System.IO.Path]::Combine($b,
			$db + "-" + ([int]$now.DayOfWeek).ToString("0") + "-" + $now.ToString("yyyy-MM-dd") + ".bak")

		BackupDB $db $b $dailyfull
		$b = Compress $db $b
		$backupfiles = $backupfiles + " " + $b
	}
}

Write-Host ("Backup End: " + [DateTime]::Now.ToString())

$total = 0
foreach($fn in $backupfiles.Trim().Split(' '))
{
	$fi = New-Object System.IO.FileInfo([System.IO.Path]::Combine($remotebackupdir, $fn))

	if ($fi.Length) {
		Write-Host ($fn + " " + $fi.Length.ToString())
		$total = $total + $fi.Length
	}
	else {
		Write-Host ($fn + " missing")
	}
}
Write-Host ("total size " + $total)

Stop-Transcript

if ($dbstatmonthly -ne $null) {
	$dbstatmonthly
}

$dbstat

if ($smtpServer -ne "")
{
	$log = Get-Content $logfile
	$bErrors = 0

	$body = New-Object System.Text.StringBuilder
	$dummy = $body.AppendLine("<html><head><style type='text/css'>* { font-family: Verdana, Arial; } body { font-size: 0.9em; } td{ font-size: 0.8em; }</style></head><body>")
	
	if ($dbstatmonthly -ne $null) {
		$dummy = $body.AppendLine("Monthly backup<p><table>")
		$dummy = $body.AppendLine("<tr valign='top'><td>Database</td><td align='center'>Status</td><td>Backup File</td><td align='right'>.bak&nbsp;kB</td><td align='right'>.zip&nbsp;kB</td></tr>")
		$sl = new-object System.collections.sortedlist($dbstatmonthly)
		$total = [long]0
		$totalzip = [long]0
		
		foreach($dbk in $sl.Keys) {
			$db = $dbstatmonthly[$dbk]

			$size = [long]0
			if ([long]::TryParse($db[3], [ref]$size)) {
				$total = $total + $size
			}
			$zipsize = [long]0
			if ([long]::TryParse($db[4], [ref]$zipsize)) {
				$totalzip = $totalzip + $zipsize
			}

			$dummy = $body.AppendLine("<tr valign='top'><td>" + $dbk + "</td><td align='center'>" + $db[0] + "</td><td>" + `
				$db[2] + "</td><td align='right'>" + `
				([long]($size / 1024)).ToString("N0") + "</td><td align='right'>" + `
				([long]($zipsize / 1024)).ToString("N0")  + "</td></tr>")
				
		}
		$dummy = $body.AppendLine("<tr valign='top'><td>Total</td><td></td><td></td><td align='right'>" + `
			([long]($total / 1024)).ToString("N0") + "</td><td align='right'>" + `
			([long]($totalzip / 1024)).ToString("N0") + "</td></tr>")
		$dummy = $body.AppendLine("</table><p>");
	}

	if ($doweekly -eq $now.DayOfWeek) {	# weekly backup
		$dummy = $body.AppendLine("Weekly backup<p><table>")
	} else {
		$dummy = $body.AppendLine("Daily backup<p><table>")
	}
	$dummy = $body.AppendLine("<tr valign='top'><td>Database</td><td align='center'>Status</td><td>Backup File</td><td align='right'>.bak&nbsp;kB</td><td align='right'>.zip&nbsp;kB</td></tr>")
	
	$sl = new-object System.collections.sortedlist($dbstat)
	$total = [long]0
	$totalzip = [long]0

	foreach($dbk in $sl.Keys) {
		$db = $dbstat[$dbk]

		$size = [long]0
		if ([long]::TryParse($db[3], [ref]$size)) {
			$total = $total + $size
		}
		$zipsize = [long]0
		if ([long]::TryParse($db[4], [ref]$zipsize)) {
			$totalzip = $totalzip + $zipsize
		}

		$dummy = $body.AppendLine("<tr valign='top'><td>" + $dbk + "</td><td align='center'>" + $db[0] + "</td><td>" + `
			$db[2] + "</td><td align='right'>" + `
			([long]($size / 1024)).ToString("N0") + "</td><td align='right'>" + `
			([long]($zipsize / 1024)).ToString("N0")  + "</td></tr>")
	}

	$dummy = $body.AppendLine("<tr valign='top'><td>Total</td><td></td><td></td><td align='right'>" + `
		([long]($total / 1024)).ToString("N0") + "</td><td align='right'>" + `
		([long]($totalzip / 1024)).ToString("N0") + "</td></tr>")
	$dummy = $body.AppendLine("</table><p><pre>");
	foreach($line in $log)
	{
		$dummy = $body.AppendLine($line.ToString().Trim())
		
		$sline = $line.ToString()
		if ($sline.Contains("missing") -or $line.Contains("failed")) {
			$bErrors = 1
		}
	}
	$dummy = $body.AppendLine("</pre>");
	
	$smtp = new-object Net.Mail.SmtpClient($smtpServer)
	if ($bErrors) {
		$subject = $subject + ": check errors"
	}
	#$smtp.Send($emailFrom, $emailTo, $subject, $body.ToString())

	$msg = New-Object Net.Mail.MailMessage($emailFrom, $emailTo, $subject, $body.ToString())
	$msg.IsBodyHTML = $true
	$smtp.Send($msg)
}
