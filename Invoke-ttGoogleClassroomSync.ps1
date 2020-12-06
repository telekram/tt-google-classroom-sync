param(
  [string]$csvpath=$null,
  [switch]$AddSubjects,
  [switch]$AddClasses,
  [switch]$AddTeachersToSubjects,
  [switch]$AddTeachersToClasses,
  [switch]$AddStudentsToClasses,
  [switch]$GetRemoteCourses
) 

$script:config = Get-Content -Raw -Path .\config.json | ConvertFrom-Json

Remove-Item -Path .\log.txt

Import-Module .\lib\dataSources.psm1 -Force -Scope Local
Import-Module .\lib\sessionManager.psm1 -Force -Scope Local

$DS = Get-DataSourceObject($csvpath)
$CA = $script:config.classroomAdmin
$AY = $script:config.academicYear
 
$session = Get-ScriptPSSession

Invoke-Command -Session $session -ScriptBlock {

  $DataSet = $Using:DS

  $academicYear = $Using:AY
  $classroomAdmin = $Using:CA
  $addSubjects = $Using:AddSubjects
  $addClasses = $Using:AddClasses
  $addTeachersToSubjects = $Using:AddTeachersToSubjects
  $addTeachersToClasses = $Using:AddTeachersToClasses
  $addStudentsToClasses = $Using:AddStudentsToClasses
  $getRemoteCourses = $Using:GetRemoteCourses
  
  function Main {

    if($getRemoteCourses){
      Get-CoursesFromGoogle
    }
    
    if($addSubjects) {
      Add-SubjectCoursesToGoogle
    }

    if($addClasses) {
      Add-ClassCoursesToGoogle
    }

    if($addTeachersToSubjects) {
      Add-TeachersToSubjects
    }

    if($addTeachersToClasses) {
      Add-TeachersToClasses
    }

    if($addStudentsToClasses){
      Add-StudentsToClasses
    }
  }


  function Get-CoursesFromGoogle {

    [System.Collections.ArrayList]$script:CloudCourses = @()

    $gCourses = gam print courses teacher $classroomAdmin | Out-String
	
    $courses = $gCourses | ConvertFrom-Csv -Delim ','
    $courses | ForEach-Object {

      
      [void]$script:CloudCourses.Add([PSCustomObject]@{
        Id = $_.id
        Name = $_.Name
        CourseState = $_.courseState
        CreationTime = $_.creationTime
        Description = $_.description
        DescriptionHeading = $_.DescriptionHeading
        Section = $_.Section
        EnrollmentCode = $_.EnrollmentCode
      })
    }

    $coursesArray = @()
    $script:CloudCourses | ForEach-Object {

      Write-Host 'doing stuff...'
      

      $courseInfo = gam info course $_.id | Out-String
      $courseInfo = $courseInfo | ConvertFrom-Csv -Delim ','

      
      $coursesArray += $courseInfo
    }
  }

  function Add-SubjectCoursesToGoogle {
    
    $progressCounter = 0

    $DataSet.Subjects | ForEach-Object {

      $subjectCode = $_.SubjectCode
      $subjectName = $_.SubjectName
      $facultyName = $_.FacultyName


      $command = "gam create course alias $subjectCode name '$subjectCode (Teachers)' section '$subjectName' heading $facultyName description 'Subject Domain: $facultyName' teacher $classroomAdmin status active"
      
      #Write-Host $command
      Invoke-Expression $command | Out-File -FilePath .\log.txt -Append
      
      $progressBarMessage = "Adding subject course: $subjectCode - $subjectName "
      $progressCounter = $progressCounter + 1

      Get-ProgressBar (
        $progressCounter,
        $DataSet.Subjects.Count,
        $progressBarMessage,
        'Magenta'
      )
    }

  }

  function Add-ClassCoursesToGoogle {
    
    $progressCounter = 0

    $DataSet.Subjects | ForEach-Object {
      
      $subjectName = $_.SubjectName
      $facultyName = $_.FacultyName

      $_.ClassCodes | ForEach-Object {

        $cc = $_
        $alias = $academicYear + $cc
        
        $command = "gam create course alias $alias name $_ section '$subjectName' heading $facultyName description 'Subject Domain: $facultyName' teacher $classroomAdmin status active"
        
        #Write-Host $command
        Invoke-Expression $command | Out-File -FilePath .\log.txt -Append
      }

      $progressBarMessage = "Adding class course: $cc"
      $progressCounter = $progressCounter + 1

      Get-ProgressBar (
        $progressCounter,
        $DataSet.Subjects.Count,
        $progressBarMessage,
        'Magenta'
      )
    }
  }

  function Add-TeachersToSubjects {

    $DataSet.Subjects
    $progressCounter = 0

    $DataSet.Subjects | ForEach-Object {

      $subjectCode = $_.SubjectCode
      $teachers = $_.Teachers
      $domainLeader = $_.DomainLeader

      if(![string]::IsNullOrWhiteSpace($domainLeader)) {
        $command = "gam course $subjectCode add teacher $domainLeader"
        #Write-Host $command
        Invoke-Expression $command | Out-File -FilePath .\log.txt -Append
      }

      $teachers | ForEach-Object {
        
        $teacher = $_
        $command = "gam course $subjectCode add teacher $teacher"

        #Write-Host $command
        Invoke-Expression $command | Out-File -FilePath .\log.txt -Append
      }  
      
      $progressBarMessage = "Adding teacher to course course: $subjectCode"
      $progressCounter = $progressCounter + 1

      Get-ProgressBar (
        $progressCounter,
        $DataSet.Subjects.Count,
        $progressBarMessage,
        'Magenta'
      )
    }
  }

  function Add-TeachersToClasses {

    $progressCounter = 0

    $DataSet.Subjects | ForEach-Object {

      $subjectCode = $_.SubjectCode
      $domainLeader = $_.DomainLeader
      $teachers = $_.Teachers
      $classCodes = $_.ClassCodes

      $classCodes | ForEach-Object {
        
        $class = $academicYear + '-' + $_

        if(![string]::IsNullOrWhiteSpace($domainLeader)) {
          $command = "gam course $class add teacher $domainLeader"
          #Write-Host $command
          Invoke-Expression $command | Out-File -FilePath .\log.txt -Append
        }


        $teachers | ForEach-Object {

          $t = $_

          if(![string]::IsNullOrWhiteSpace($t)) {
            $command = "gam course $class add teacher $t"
            #Write-Host $command
            Invoke-Expression $command | Out-File -FilePath .\log.txt -Append
          }
        }
      }  

      $progressBarMessage = "Adding teacher to classes in subject course: $subjectCode"
      $progressCounter = $progressCounter + 1

      Get-ProgressBar (
        $progressCounter,
        $DataSet.Subjects.Count,
        $progressBarMessage,
        'Magenta'
      )
    }
  }


  function Add-StudentsToClasses {

    $progressCounter = 0

    $DataSet.Classes | ForEach-Object {

      $class = $academicYear + '-' + $_.ClassCode
      
      $students = $_.StudentCodes

      $students | ForEach-Object {
        $s = $_
        $command = "gam course $class add student $s"
        #Write-Host $command
        Invoke-Expression $command | Out-File -FilePath .\log.txt -Append
      }

      $progressBarMessage = "Adding students to course: $class"
      $progressCounter = $progressCounter + 1

      Get-ProgressBar (
        $progressCounter,
        $DataSet.Classes.Count,
        $progressBarMessage,
        'Magenta'
      )
    }
  }

  function Get-ProgressBar ($arg) {    
    
    $progressCounter = $arg[0]
    $totalCount = $arg[1]
    $progressBarMessage = $arg[2]
    $progressBarColor = $arg[3]

    #$Host.PrivateData.ProgressBackgroundColor=$progressBarColor

    Write-Progress -Activity $progressBarMessage -Status "Progress:" -PercentComplete ($progressCounter / $totalCount * 100)
  }

  Main

} 

Clear-ScriptPSSession
