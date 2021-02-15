$script:config = Get-Content -Raw -Path .\config.json | ConvertFrom-Json
[PSCustomObject]$script:TimetableObj = @{}
[PSCustomObject]$script:Dataset = @{}

function Get-DataSourceObject ($csvPath) {

  function Main {
    
    Get-TimetableDataAsObjectFromCsvFiles
    Add-SubjectsToDataset
    Get-CompositeClasses
    Add-ClassCodesToSubjectsObject
    #Add-TeachersToSubjectObject
    #Add-DomainLeadersToSubjectObject
    #Add-StudentsToDataset
    #Add-ClassCodesToStudents
    # Add-ClassCodesToDataset
    # Add-StudentsToClassCodes

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
    
    $indexesOfSubjectsWihtoutAnyClasses = @()
    $index = 0

    $script:Dataset.Subjects | ForEach-Object {
      if($_.ClassCodes.Count -eq 0) {
        $indexesOfSubjectsWihtoutAnyClasses += $index
      }
      $index = $index + 1
    }
    
    $indexesOfSubjectsWihtoutAnyClasses | 
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

      if(Test-ActiveDirectoryForUser($teacherCode)) {

        $compositeClassCandiateRows = $script:TimetableObj.Timetable | 
          Sort-Object -Property 'Class Code' | 
            Where-Object {
              $_.'Day No' -eq $dayNumber -and 
              $_.'Period No' -eq $periodNumber -and 
              $_.'Teacher Code' -eq $teacherCode
            }
      } else {
        "Staff code '$teacherCode' not found in Active Directory" | Out-File -FilePath .\log.txt -Append
      }

    

      if ($compositeClassCandiateRows.Count -gt 1) { # if theres more than 1 it's a composite class

        if(!$alreadyProcessedClasscodes.Contains($classCode )) { # prevent duplicates    

          $compositeClassName = $compositeClassCandiateRows.'Class Code'[0] +  '-' + $compositeClassCandiateRows.'Class Code'[1]

          $compositeClasses += [PSCustomObject]@{
            'SubjectCode' = $compositeClassName
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
            $teachers += $teacherCode.ToLower() + $script:config.domainName
          } else {
            "Staff code '$teacherCode' not found in Active Directory" | Out-File -FilePath .\log.txt -Append
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

            if(Test-ActiveDirectoryForUser($_.Member)) {
              $subject | Add-Member -MemberType NoteProperty -Name DomainLeader -Value $dLeader
            } else {
              "Domain Leader: '$dLeader' not found in Active Directory" | Out-File -FilePath .\log.txt -Append
            }
          }     
        }
      }

      $progressCounter = $progressCounter + 1
      $progressBarMessage = 'Adding Domain Leader to subject: ' + $subject.SubjectCode

      Get-ProgressBar (
        $progressCounter,
        $script:Dataset.Subjects.Count,
        $progressBarMessage,
        'Green'
      )
    }
  }

  function Add-StudentsToDataset {

    [System.Collections.ArrayList]$students = @()

    $script:TimetableObj.StudentLessons | 
      Sort-Object -Property 'Student Code' -Unique |
        ForEach-Object {

          $studentCode = $_.'Student Code'.ToLower() + $script:config.domainName

          if(Test-ActiveDirectoryForUser($studentCode)){

            [void]$students.Add([PSCustomObject]@{ 
              StudentCode = $studentCode
            })

          }
        }
    $script:Dataset.Students = $students
  }


  function Add-ClassCodesToStudents {

    $progressCounter = 0

    $script:Dataset.Students | ForEach-Object {
    
      [Array]$cCodes = @()

      $studentCode = $_.StudentCode

      $studentLessonRows = $script:TimetableObj.StudentLessons | Where-Object {
        $studentLessonsStudentCode = $_.'Student Code'.ToLower() + $script:config.domainName

        $studentLessonsStudentCode -eq $studentCode
      }

      $studentLessonRows | ForEach-Object {
        $cCodes += $_.'Class Code'
      }

      
      $_ | Add-Member -MemberType NoteProperty -Name ClassCodes -Value $cCodes

      $progressCounter = $progressCounter + 1
      $progressBarMessage = 'Adding classes for: ' + $studentCode.ToLower() + ' [' + $cCodes + ']'

      Get-ProgressBar (
        $progressCounter,
        $script:Dataset.Students.count,
        $progressBarMessage,
        'DarkCyan'
      )
    }
  }

  function Add-ClassCodesToDataset {
    
    [System.Collections.ArrayList]$classes = @()

    $script:TimetableObj.StudentLessons | 
      Sort-Object -Property 'Class Code' -Unique |
        ForEach-Object {

          [void]$classes.Add([PSCustomObject]@{ 
            ClassCode = $_.'Class Code'
          })
        }

    $script:Dataset.Classes = $classes
    
  }

  function Add-StudentsToClassCodes {

    $progressCounter = 0

    $script:Dataset.Classes | ForEach-Object {

      [Array]$students = @()
      [Array]$studentsInActiveDirectory = @()

      $classCode = $_.ClassCode

      $studentLessonRows = $script:TimetableObj.StudentLessons | Where-Object {
        $_.'Class Code' -eq $classCode
      }

      $studentLessonRows | Sort-Object -Unique | ForEach-Object {
        $students += $studentLessonRows.'Student Code'
      }

      $students | ForEach-Object {

        if (Test-ActiveDirectoryForUser($_)) {
          $studentsInActiveDirectory += $_.ToLower() + $script:config.domainName
        } else {
          "Student: '$_' not found in Active Directory" | Out-File -FilePath .\log.txt -Append
        }
      }

      $_ | Add-Member -MemberType NoteProperty -Name StudentCodes -Value $studentsInActiveDirectory

      $progressCounter = $progressCounter + 1
      $progressBarMessage = 'Adding students to: ' + $classCode + ' [' + $studentsInActiveDirectory + ']'

      Get-ProgressBar (
        $progressCounter,
        $script:Dataset.Classes.Count,
        $progressBarMessage,
        'Magenta'
      )
    }
  }

  function Test-ActiveDirectoryForUser($user) {

    if($script:config.isActiveDirectoryAvailable) { 

      $searchBase = "OU=Users,OU=CSC,DC=curric,DC=cheltenham-sc,DC=wan"
      return Get-ADUser -Filter { SamAccountName -eq $user } -SearchBase $searchBase

    } else {
      return 1
    }

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
  return $script:Dataset
}



