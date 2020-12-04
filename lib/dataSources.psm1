#$script:config = Get-Content -Raw -Path .\config.json | ConvertFrom-Json
[PSCustomObject]$script:TimetableObj = @{}
[PSCustomObject]$script:Dataset = @{}

function Get-DataSourceObject ($csvPath) {
  function Main {
    
    Get-TimetableDataAsObjectFromCsvFiles
    Add-SubjectsObject
    Get-CompositeClasses
    Add-ClassCodesToSubjectsObject
    Add-TeachersToSubjectObject
    Add-DomainLeadersToSubject
    Get-StudentsObjectFromTimetableObject
    Add-ClassCodesToStudents
    
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


  function Add-SubjectsObject {

    [System.Collections.ArrayList]$subjects = @()

    $progressCounter = 0
    $prgressTotal = $script:TimetableObj.ClassNames | Sort-Object -Property 'Subject Code' -Unique

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
            $progressBarMessage, 
            'Magenta'
          )
        }
    
    $script:Dataset.Subjects = $subjects
  }

  
  function Add-ClassCodesToSubjectsObject {

    $compositeClassList = $script:Dataset.CompositeClassList
    
    $progressCounter = 0
    
    $script:Dataset.Subjects | ForEach-Object {
  
      $sc = $_.SubjectCode

      $cc = $script:TimetableObj.ClassNames | 
        Where-Object { 
          $_.'Subject Code' -eq $sc
        }
  
      
      $DiffOfCompositeAndStandardClasses = Compare-Object -ReferenceObject $cc.'Class Code' -DifferenceObject $compositeClassList

      if($DiffOfCompositeAndStandardClasses.SideIndicator -eq '<='){
        $_ | Add-Member -MemberType NoteProperty -Name ClassCodes -Value $cc.'Class Code'
      } 

      $progressCounter = $progressCounter + 1

      $c = $cc.'Class Code'
      $progressBarMessage = "Adding class: $c to Subjects Object"

      Get-ProgressBar ($progressCounter, 
        $script:Dataset.Subjects.count, 
        $progressBarMessage, 
        'DarkCyan'
      )
    }
    
    $subjectsWithoutClassesIndex = @()
    $index = 0
    $script:Dataset.Subjects | ForEach-Object {
      if($_.ClassCodes.Count -eq 0) {
        $subjectsWithoutClassesIndex += $index
      }
      $index = $index + 1
    }
    
    $subjectsWithoutClassesIndex | Sort-Object -Descending | ForEach-Object {
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


      if(Test-ActiveDirectoryForUser($teacherCode)) {

        $compositeClassCandiateRows = $script:TimetableObj.Timetable | 
          Sort-Object -Property 'Class Code' | 
            Where-Object {
              $_.'Day No' -eq $dayNumber -and 
              $_.'Period No' -eq $periodNumber -and 
              $_.'Teacher Code' -eq $teacherCode
            }

      } else {
        Write-Warning "Staff code '$teacherCode' not found in Staff OU"
      }
    

      if ($compositeClassCandiateRows.Count -gt 1) { # if theres more than 1 it's a composite class

        if(!$alreadyProcessedClasscodes.Contains($classCode )) { # prevent duplicates    

          $compositeClassName = $compositeClassCandiateRows.'Class Code'[0] +  '-' + $compositeClassCandiateRows.'Class Code'[1]

          $compositeClasses += [PSCustomObject]@{
            'SubjectName' = "Composite $compositeClassName"  
            'ClassCodes' = $compositeClassCandiateRows.'Class Code'
            'Teachers' = $compositeClassCandiateRows.'Teacher Code' | Sort-Object -Unique
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
  }

  function Add-TeachersToSubjectObject {

    $progressCounter = 0

    $script:Dataset.Subjects | ForEach-Object {
    
      [Array]$teachers = @()
      
      $subject = $_
      $subjectCode = $subject.SubjectCode
  
      $subject.ClassCodes | ForEach-Object {
  
        $cCodes = $_
  
        $t = $TimetableObj.Timetable | Where-Object { 
          $_.'Class Code' -eq $cCodes 
        }
          
        $allTimetabledTeacherCodes = $t.'Teacher Code' | Where-Object { $_ }

        $allTimetabledTeacherCodes | ForEach-Object {

          $teacherCode = $_
          
          if(Test-ActiveDirectoryForUser($teacherCode)) {
            $teachers += $teacherCode
          } else {
            Write-Warning "Function: Add-TeachersToSubjectObject - Staff code $teacherCode not found in Staff OU"
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
        $progressBarMessage,
        'Magenta'
      )
    }
  }
  

  function Add-DomainLeadersToSubject {

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
  
          $isDomainLeaderAlreadyAdded = [bool]($subject.PSObject.Properties.Name -match 'DomainLeaders')

          if(!$isDomainLeaderAlreadyAdded) {
            $subject | Add-Member -MemberType NoteProperty -Name DomainLeaders -Value $_.Member
          }     
        }
      }

      $progressCounter = $progressCounter + 1
      $progressBarMessage = 'Adding Domain Leader [' + $_.Member + '] to subject: ' + $subject.SubjectCode

      Get-ProgressBar (
        $progressCounter,
        $script:Dataset.Subjects.Count,
        $progressBarMessage,
        'Green'
      )
    }
  }

  function Get-StudentsObjectFromTimetableObject {

    [System.Collections.ArrayList]$students = @()

    $script:TimetableObj.StudentLessons | 
      Sort-Object -Property 'Student Code' -Unique |
        ForEach-Object {

          [void]$students.Add([PSCustomObject]@{ 
            StudentCode = $_.'Student Code'
          })
        }

    $script:Dataset.Students = $students
  }


  function Add-ClassCodesToStudents {

    $progressCounter = 0

    $script:Dataset.Students | ForEach-Object {
    
      [Array]$cCodes = @()

      $studentCode = $_.StudentCode

      $studentLessonRows = $script:TimetableObj.StudentLessons | Where-Object {
        $_.'Student Code' -eq $studentCode
      }

      $studentLessonRows | ForEach-Object {
        $cCodes += $_.'Class Code'
      }

      
      $_ | Add-Member -MemberType NoteProperty -Name ClassCodes -Value $cCodes

      $progressCounter = $progressCounter + 1
      $progressBarMessage = 'Adding classes for: ' + $studentCode + ' [' + $cCodes + ']'

      Get-ProgressBar (
        $progressCounter,
        $script:Dataset.Students.count,
        $progressBarMessage,
        'DarkCyan'
      )
    }
  }


  function Test-ActiveDirectoryForUser($user) {
    return Get-ADUser -Filter { SamAccountName -eq $user } -SearchBase "OU=Staff,OU=Users,OU=CSC,DC=curric,DC=cheltenham-sc,DC=wan"
  }

  
  function Get-ProgressBar ($arg) {    
    
    $progressCounter = $arg[0]
    $totalCount = $arg[1]
    $progressBarMessage = $arg[2]
    $progressBarColor = $arg[3]

    $Host.PrivateData.ProgressBackgroundColor=$progressBarColor

    Write-Progress -Activity $progressBarMessage -Status "Progress:" -PercentComplete ($progressCounter / $totalCount * 100)
  }

  Main
  $script:Dataset.Subjects | Format-Table
  $script:Dataset.Students  | Format-Table
}



