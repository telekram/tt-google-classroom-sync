#$script:config = Get-Content -Raw -Path .\config.json | ConvertFrom-Json

function Get-DataSourceObject () {
  function Main {

    Get-TimetableDataAsObjectFromCsvFiles
    Get-SubjectsObjectForDataSource
    Get-ClassCodesForSubjects
    Get-TeachersForSubjectsObject

  }


  function Get-TimetableDataAsObjectFromCsvFiles {

    $classNamesCsv = Import-Csv -Path '.\tt\Class Names.csv'
    $studentLessonsCsv = Import-Csv -Path '.\tt\Student Lessons.csv'
    $timeTableCsv = Import-Csv -Path '.\tt\Timetable.csv'
  
    $timetableData = [ordered]@{
      ClassNames = $classNamesCsv;
      StudentLessons = $studentLessonsCsv
      Timetable = $timeTableCsv
    }
  
    $script:timetableObj = New-Object -Type PSObject -Property $timetableData

  }


  function Get-SubjectsObjectForDataSource {

    $progressCounter = 0

    $script:subjects = @()

    $script:timetableObj.ClassNames |
      Sort-Object -Property 'Subject Code' -Unique | 
        ForEach-Object {  
  
          $faculty = $_.'Faculty Name'.Split('_')
  
          $script:subjects += [PSCustomObject]@{
            SubjectCode = $_.'Subject Code';
            SubjectName = $_.'Subject Name';
            FacultyName = $faculty[0]
          }
  
          $progressCounter = $progressCounter + 1

          Get-ProgressBar (
            $progressCounter, 
            $script:timetableObj.ClassNames.count, 
            'Generating Subjects', 
            'Magenta'
          )
        } 
  }

  
  function Get-ClassCodesForSubjects {

    $progressCounter = 0

    $script:subjects | ForEach-Object {
  
      $sc = $_.SubjectCode

      $cc = $script:timetableObj.ClassNames | 
        Where-Object { 
          $_.'Subject Code' -eq $sc
        }
  
      $_ | Add-Member -MemberType NoteProperty -Name ClassCodes -Value $cc.'Class Code'
      
      $progressCounter = $progressCounter + 1

      $c = $cc.'Class Code'
      $progressBarMessage = "Adding class: $c to Subjects Object"

      Get-ProgressBar ($progressCounter, 
        $subjects.count, 
        $progressBarMessage, 
        'Cyan'
      )
  
    }
  }
  

  function Get-TeachersForSubjectsObject {

    $progressCounter = 0

    $script:subjects | ForEach-Object {
    
      $subj = $_
      $teachers = @()
      $code = $_.SubjectCode
  
      $_.ClassCodes | ForEach-Object {
  
        $cc = $_
  
        $t = $timetableObj.Timetable | Where-Object { 
          $_.'Class Code' -eq $cc 
        }
  
        if ($teachers -notcontains $t.'Teacher Code') {
          $teachers += $t.'Teacher Code'
        }
      }
  
        
      $uniqueTeachers = $teachers | Select-Object -Unique

      $subj | Add-Member -MemberType NoteProperty -Name Teachers -Value $uniqueTeachers

      $progressCounter = $progressCounter + 1
      $progressBarMessage = "Adding teacher(s): [ $uniqueTeachers ] to subject: $code"

      Get-ProgressBar (
        $progressCounter,
        $subjects.count,
        $progressBarMessage,
        'Yellow'
      )
    }
  }

  function Get-ProgressBar ($a) {    
    
    $i = $a[0]
    $x = $a[1]
    $y = $a[2]
    $z = $a[3]

    $Host.PrivateData.ProgressBackgroundColor=$z 

    Write-Progress -Activity $y -Status "Progress:" -PercentComplete ($i / $x *100)
  }

  Main
  #$script:subjects
}



