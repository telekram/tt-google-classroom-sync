$script:config = Get-Content -Raw -Path .\config.json | ConvertFrom-Json
[PSCustomObject]$script:TimetableObj = @{}
[PSCustomObject]$script:Dataset = @{}

function Get-DataSourceObject ($csvPath) {

  function Main {
    Clear-Log
    Approve-ScriptConditions
    Get-TimetableDataAsObjectFromCsvFiles
    Add-SubjectsToDataset
    Get-CompositeClasses
    Add-ClassCodesToSubjectsObject
    Add-TeachersToSubjectObject
    add-DomainLeadersToSubjectObject
    Add-StudentsToClassCodes
  }



  function Get-TimetableDataAsObjectFromCsvFiles {

    $classNamesCsv = Import-Csv -Path "$csvPath\Class Names.csv"
    $studentLessonsCsv = Import-Csv -Path "$csvPath\Student Lessons.csv"
    $timeTableCsv = Import-Csv -Path "$csvPath\Timetable.csv"
    $unschedulesDutiesCsv = Import-Csv -Path "$csvPath\Unscheduled Duties.csv"
  
    $timetableData = [hashtable]@{
      ClassNames = $classNamesCsv
      StudentLessons = $studentLessonsCsv
      Timetable = $timeTableCsv
      UnscheduledDuties = $unschedulesDutiesCsv
    }
  
    $script:TimetableObj = $timetableData
  }


  function Add-SubjectsToDataset {

    [System.Collections.ArrayList]$subjects = @()

    $progressCounter = 0
    
    $prgressTotal = $script:TimetableObj.ClassNames | 
      Sort-Object -Property 'Subject Code' -Unique

    $script:TimetableObj.ClassNames |
      Sort-Object -Property 'Subject Code' -Unique | 
        ForEach-Object {  
  
          $facultyParts = $_.'Faculty Name'.Split('_')
  
          $s = [PSCustomObject]@{
            SubjectCode = $_.'Subject Code'
            SubjectName = $_.'Subject Name'
            FacultyName = $facultyParts[0]
          }

          [void]$subjects.Add($s)
  
          $progressCounter = $progressCounter + 1

          $progressBarMessage = "Adding subject: " + $_.'Subject Code'

          Get-ProgressBar (
            $progressCounter, 
            $prgressTotal.count, 
            $progressBarMessage
          )
        }
    
    $script:Dataset.Subjects = $subjects
    Update-Log("Dataset: subjects added")
  }

  
  function Add-ClassCodesToSubjectsObject {
    
    $progressCounter = 0
    
    $script:Dataset.Subjects | ForEach-Object {
  
      $subject = $_
      $sc = $_.SubjectCode


      $cc = $script:TimetableObj.ClassNames | 
        Where-Object { 
          $_.'Subject Code' -eq $sc
        }

      [System.Collections.ArrayList]$classCodes = @()

      $cc.'Class Code' | ForEach-Object {
        $classCode = $_

        if ($script:Dataset.CompositeClassList -notcontains $classCode) {
          [void]$classCodes.Add([PSCustomObject]@{
            Class = $classCode
          })
        }
      }

      $subject | Add-Member -MemberType NoteProperty -Name ClassCodes -Value $classCodes

      $progressCounter = $progressCounter + 1

      $c = $cc.'Class Code'
      $progressBarMessage = "Adding class: $c to Subjects Object"

      Get-ProgressBar ($progressCounter, 
        $script:Dataset.Subjects.count, 
        $progressBarMessage
      )
    }
    
    Remove-SubjectsWithoutClasses
    Update-Log("Dataset: classcodes added to subjects")
  }

  function Remove-SubjectsWithoutClasses {
    [array]$SubjectsWihtoutClassesIndexes = @()
    $index = 0

    $script:Dataset.Subjects | ForEach-Object {
      if($_.ClassCodes.Count -eq 0) {
        $SubjectsWihtoutClassesIndexes += $index
      }
      $index = $index + 1
    }
    
    $SubjectsWihtoutClassesIndexes | 
      Sort-Object -Descending | 
        ForEach-Object {
          $script:Dataset.Subjects.RemoveAt($_)
    }
  }
  


  function Get-CompositeClasses {
    
    $compositeClasses = @()
    $alreadyProcessedClasscodes = @()

    $progressCounter = 0
    
    $uniqueSortedClassCodes = $script:TimetableObj.Timetable | 
      Sort-Object -Property 'Class Code' -Unique

    $uniqueSortedClassCodes | ForEach-Object {

      $classCode = $_.'Class Code'
    
      $dayNumber = $_.'Day No'
      $periodNumber = $_.'Period No'
      $teacherCode = $_.'Teacher Code'


      $compositeClassCandiateRows = $script:TimetableObj.Timetable | 
        Sort-Object -Property 'Class Code' | 
          Where-Object {
            $_.'Day No' -eq $dayNumber -and 
            $_.'Period No' -eq $periodNumber -and 
            $_.'Teacher Code' -eq $teacherCode
          }


      if ($compositeClassCandiateRows.Count -gt 1) { # if theres more than 1 it's a composite class

        if(!$alreadyProcessedClasscodes.Contains($classCode )) { # prevent duplicates    

          $compositeClassName = $compositeClassCandiateRows.'Class Code'[0] +  '-' + $compositeClassCandiateRows.'Class Code'[1]

          $teachers = @()

          $compositeClassCandiateRows.'Teacher Code' | Sort-Object -Unique | ForEach-Object {
            $teacher = $_.ToLower() + $script:config.domainName
            $teachers += $teacher
          }
          

          $compositeClasses += [PSCustomObject]@{
            'SubjectCode' = $compositeClassName
            'SubjectName' = "Composite $compositeClassName"  
            'ClassCodes' = $compositeClassCandiateRows.'Class Code'
            'Teachers' = $teachers
          }
        }

        $compositeClassCandiateRows.'Class Code' | ForEach-Object {
          $alreadyProcessedClasscodes += $_
        }      
      }

      $progressCounter = $progressCounter + 1

      Get-ProgressBar ($progressCounter, 
        $uniqueSortedClassCodes.count, 
        'Searching for composite classes', 
        'DarkCyan'
      )
      
    }
    $compositeClassList = $alreadyProcessedClasscodes | Sort-Object | Get-Unique

    $script:Dataset.CompositeClasses = $compositeClasses
    $script:Dataset.CompositeClassList = $compositeClassList

    Update-Log("Dataset: composite classes added")
  }

  function Add-TeachersToSubjectObject {

    $progressCounter = 0

    $script:Dataset.Subjects | ForEach-Object {
    
      [Array]$teachers = @()
      
      $subject = $_
      $subjectCode = $subject.SubjectCode
  
      $subject.ClassCodes | ForEach-Object {
  
        $cCodes = $_.Class
  
        $t = $TimetableObj.Timetable | Where-Object { 
          $_.'Class Code' -eq $cCodes 
        }
          
        $allTimetabledTeacherCodes = $t.'Teacher Code' | Where-Object { $_ }

        $allTimetabledTeacherCodes | ForEach-Object {

          $teacherCode = $_
          
          if(Test-ActiveDirectoryForUser($teacherCode)) {
            $teachers += $teacherCode.ToLower() + $script:config.domainName
          } else {
            Update-Log("f:Add-TeachersToSubjectObject --> staff code '$teacherCode' not found in Active Directory")
          }
        }
      }
        
      $uniqueTeachers = $teachers | Select-Object -Unique

      $subject | Add-Member -MemberType NoteProperty -Name Teachers -Value $uniqueTeachers

      $progressCounter = $progressCounter + 1
      $progressBarMessage = "Adding teacher(s): [ $uniqueTeachers ] to subject: $subjectCode"

      Get-ProgressBar (
        $progressCounter,
        $script:Dataset.Subjects.count,
        $progressBarMessage
      )
    }
    Update-Log("Dataset: teachers added to subject")
  }
  

  function Add-DomainLeadersToSubjectObject {

    $domainLeaders = @()

    $script:TimetableObj.UnscheduledDuties | Where-Object {
      $_.'Duty Name' -like '*Domain Leader*'
    } | ForEach-Object {

      $domainNameParts = $_.'Duty Name'.Split('_')
      $domainMember = $_.'Teacher Code'

      $domainLeaders += [PSCustomObject]@{
        'Domain' = $domainNameParts[1]
        'Member' = $domainMember
      }
    }

    $progressCounter = 0

    $script:Dataset.Subjects | ForEach-Object {

      $subject = $_
      $faculty = $_.FacultyName

      $domainLeaders | Where-Object{

        if ($faculty -eq $_.Domain) {
  
          $isDomainLeaderAlreadyAdded = [bool]($subject.PSObject.Properties.Name -match 'DomainLeader')

          $dLeader = $_.Member.ToLower() + $script:config.domainName

          if(!$isDomainLeaderAlreadyAdded) {

            if (Test-ActiveDirectoryForUser($_.Member)) {
              $subject | Add-Member -MemberType NoteProperty -Name DomainLeader -Value $dLeader
            } else {
              Update-Log("f:Add-DomainLeadersToSubjectObject --> domain leader: '$dLeader' not found in Active Directory")
            }
          }     
        }
      }

      $progressCounter = $progressCounter + 1
      $progressBarMessage = 'Adding Domain Leader to subject: ' + $subject.SubjectCode

      Get-ProgressBar (
        $progressCounter,
        $script:Dataset.Subjects.Count,
        $progressBarMessage
      )
    }
    Update-Log("Dataset: domain leaders added to subjects")
  }


  function Add-StudentsToClassCodes {

    $progressCounter = 0

    $script:Dataset.Subjects.ClassCodes | ForEach-Object { 

      [Array]$students = @()
      [Array]$studentsInActiveDirectory = @()

      $class = $_.Class

      $studentLessonRows = $script:TimetableObj.StudentLessons | Where-Object {
        $_.'Class Code' -eq $class
      }

      $studentLessonRows | Sort-Object -Unique | ForEach-Object {
        $students += $studentLessonRows.'Student Code'
      }

      $students | ForEach-Object {
        $student = $_

        if (Test-ActiveDirectoryForUser($student)) {
          $studentsInActiveDirectory += $student.ToLower() + $script:config.domainName
        } else {
          Update-Log("f:Add-StudentsToClassCodes -- > student: $student not found in Active Directory")
        }
      }

      $_ | Add-Member -MemberType NoteProperty -Name Students -Value $studentsInActiveDirectory

      $progressCounter = $progressCounter + 1
      $progressBarMessage = 'Adding students to: ' + $class + ' [' + $studentsInActiveDirectory + ']'

      Get-ProgressBar (
        $progressCounter,
        $script:Dataset.Subjects.ClassCodes.Class.Count,
        $progressBarMessage
      )
    }
    Update-Log("Dataset: students added to classcodes")
  }

  function Test-ActiveDirectoryForUser($user) {

    if($script:config.isActiveDirectoryAvailable) { 

      $searchBase = $script:config.AdSearchBase
      return Get-ADUser -Filter { SamAccountName -eq $user } -SearchBase $searchBase

    } else {
      return 1
    }
  }

  function Clear-Log {
    $file = '.\log.txt'
    if (Test-Path $file -PathType leaf) {
        Remove-Item $file
    }
  }

  function Approve-ScriptConditions {
    if (!$script:config.isActiveDirectoryAvailable) {
      Write-Warning "Script is running without Active Directory and won't be able to verify that users in timetable exist in AD."
      $answer = Read-Host "isAvtiveDirectory flag set to false in config.json. Do you wish to continue? (y/n)"
      
      if ($answer.ToLower() -ne 'y') {
        exit
      } 

      Update-Log("Dataset: generating without Active Direcoty and may contain invalid users")
    }
  }

  function Update-Log($text) {
    $dt = Get-Date -Format "dd/MM/yyyy hh:mm tt"
    $dt + ": " + $text | Out-File -FilePath '.\log.txt' -Append
  }
  
  function Get-ProgressBar ($arg) {    
    
    $progressCounter = $arg[0]
    $totalCount = $arg[1]
    $progressBarMessage = $arg[2]

    Write-Progress -Activity $progressBarMessage -Status "Progress:" -PercentComplete ($progressCounter / $totalCount * 100)
  }

  Main

  return $script:Dataset
  Update-Log("Dataset: complete")
}



