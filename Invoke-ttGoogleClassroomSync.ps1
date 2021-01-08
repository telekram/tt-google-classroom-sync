param(
  [string]$CsvPath=$null,
  [string]$ShowSubjects=$null,
  [string]$AddSubjects=$null,
  [string]$AddClasses=$null,
  [string]$AddTeachersToSubjects=$null,
  [switch]$AddTeachersToClasses,
  [switch]$AddStudentsToClasses,
  [switch]$AddCompositeClasses,
  [switch]$GetRemoteCourses
) 

$script:config = Get-Content -Raw -Path .\config.json | ConvertFrom-Json

Import-Module .\lib\dataSources.psm1 -Force -Scope Local
Import-Module .\lib\sessionManager.psm1 -Force -Scope Local

$DS = Get-DataSourceObject($csvpath)

$CA = $script:config.classroomAdmin
$AY = $script:config.academicYear
 
$session = Get-ScriptPSSession

Invoke-Command -Session $session -ScriptBlock {

  $DataSet = $Using:DS
  
  $DataSet.Subjects

  $academicYear = $Using:AY
  $classroomAdmin = $Using:CA
  $showSubjects = $Using:ShowSubjects
  $addSubjects = $Using:AddSubjects
  $addClasses = $Using:AddClasses
  $addTeachersToSubjects = $Using:AddTeachersToSubjects
  $addTeachersToClasses = $Using:AddTeachersToClasses
  $addStudentsToClasses = $Using:AddStudentsToClasses
  $addCompositeClasses = $Using:AddCompositeClasses
  $getRemoteCourses = $Using:GetRemoteCourses

  
  function Main {

    if($getRemoteCourses){
      Get-CoursesFromGoogle
    }

    if(!$null -eq $showSubjects) {
      Show-Subject($showSubjects)
    } 

    if(!$null -eq $addSubjects) {
      Add-SujectCoursesToGoogle($addSubjects)
    } 

    if(!$null -eq $addClasses) {
      Add-ClassCoursesToGoogle($addClasses)
    }

    if(!$null -eq $addTeachersToSubjects) {
      Add-TeachersToSubjects($addTeachersToSubjects)
    }

    if($addTeachersToClasses) {
      Add-TeachersToClasses
    }

    if($addStudentsToClasses){
      Add-StudentsToClasses
    }

    if($addCompositeClasses){
      Add-CompositeClasses
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

      $courseInfo = gam info course $_.id | Out-String
      $courseInfo = $courseInfo | ConvertFrom-Csv -Delim ','
      
      $coursesArray += $courseInfo
    }
  }


  function Show-Subject($subject) {

    $s = $DataSet.Subjects | 
      Where-Object { $_.SubjectCode -like "*$subject*" } 

    if(!$s) {
      Write-Host "Subject(s): '$subject' not found"
      exit
    }

    $s
  }

  function Add-SujectCoursesToGoogle($subject) {
    
    $s = $DataSet.Subjects | 
      Where-Object { $_.SubjectCode -like "*$subject*" } 

    if(!$s) {
      Write-Host "Subject(s): '$subject' not found"
      exit
    }

    $progressCounter = 0

    $s | ForEach-Object {

      $course = [PSCustomObject]@{
        Type = 'Subject'
        Code = $_.SubjectCode
        Name = $_.SubjectName
        Faculty = $_.FacultyName
      }

      Publish-Course($course)
      
      $subjectCode = $_.SubjectCode
      $subjectName = $_.SubjectName

      $progressBarMessage = "Adding subject course: $subjectCode - $subjectName"
      $progressCounter = $progressCounter + 1

      Get-ProgressBar (
        $progressCounter,
        $DataSet.Subjects.Count,
        $progressBarMessage,
        'Magenta'
      )
    }
  }

  function Publish-Course($courseAttributes) {

    $c = $courseAttributes

    if([string]::IsNullOrWhiteSpace($c.Type)){
      Write-Host "Publish course error: course type not defined. Script will exit"
      exit
    }

    if ($c.Type -eq 'Subject') {

      $alias = $c.Code
      $name = $c.Code + ' (Teachers)'
      $section = $c.Name -Replace '[^a-zA-Z0-9-_ ]', ''

    } elseif ($c.Type -eq 'Class') {

      $alias = $academicYear + '-' + $c.Code
      $name = $c.Code
      $section = $c.Name -Replace '[^a-zA-Z0-9-_ ]', ''
      $section = $academicYear + ' ' + $section 
    }
    
    $section = $c.Name -Replace '[^a-zA-Z0-9-_ ]', ''
    $description = 'Subject Domain: ' + $c.Faculty + ' - ' + $section   

    $cmd = "gam create course alias $alias name '$name' section '$section' description '$description' heading $alias teacher $classroomAdmin status active"
    Write-Host $cmd
  }


  function Add-ClassCoursesToGoogle($subject) {

    $progressCounter = 0

    $s = $DataSet.Subjects | 
      Where-Object { $_.SubjectCode -like "*$subject*" } 

    if(!$s) {
      Write-Host "Subject(s): '$subject' not found"
      exit
    }

    $s | ForEach-Object {
      
      $subjectName = $_.SubjectName
      $faculty = $_.FacultyName

      $_.ClassCodes | ForEach-Object {

        $cc = $_

        $course = [PSCustomObject]@{
          Type = 'Class'
          Code = $cc
          Name = $subjectName
          Faculty = $faculty
        }
        
        Publish-Course($course)
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

  function Add-TeachersToSubjects($subject) {

    $progressCounter = 0

    $s = $DataSet.Subjects | 
      Where-Object { $_.SubjectCode -like "*$subject*" } 

    if(!$s) {
      Write-Host "Subject(s): '$subject' not found"
      exit
    }

    $s | ForEach-Object {

      $subjectCode = $_.SubjectCode
      $teachers = $_.Teachers
      $domainLeader = $_.DomainLeader

      if(![string]::IsNullOrWhiteSpace($domainLeader)) {
        $command = "gam course $subjectCode add teacher $domainLeader"
        Invoke-Expression $command
      }

      $teachers | ForEach-Object {
        
        $teacher = $_
        $command = "gam course $subjectCode add teacher $teacher"

        Invoke-Expression $command 
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

    
  function Add-CompositeClasses {

    $progressCounter = 0

    $DataSet.CompositeClasses | ForEach-Object {

      $subjectCode = $_.SubjectCode
      $subjectName = $_.SubjectName

      $classAlias = "$academicYear-$subjectCode"

      $course = [PSCustomObject]@{
        Type = 'Class'
        Code = $_.SubjectCode
        Name = $_.SubjectName
        Faculty = ''
      }

      Publish-Course($course)

      $_.Teachers | ForEach-Object {

        $teacher = $_

        $command = "gam course $classAlias add teacher $teacher"
        Invoke-Expression $command 
      }


      $_.ClassCodes | ForEach-Object {
        
        $classCode = $_
        
        $DataSet.Classes.$classCode | ForEach-Object {

          $student = $_

          $command = "gam course $classAlias add student $student"
          Invoke-Expression $command 
        }
      }

      $progressBarMessage = "Adding subject course: $subjectCode - $subjectName"
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

    $s = $DataSet.Subjects | 
      Where-Object { $_.SubjectCode -like "*$subject*" } 

    if(!$s) {
      Write-Host "Subject(s): '$subject' not found"
      exit
    }

    $s | ForEach-Object {

      $subjectCode = $_.SubjectCode
      $domainLeader = $_.DomainLeader
      $teachers = $_.Teachers
      $classCodes = $_.ClassCodes

      $classCodes | ForEach-Object {
        
        $class = $academicYear + '-' + $_

        if(![string]::IsNullOrWhiteSpace($domainLeader)) {
          $command = "gam course $class add teacher $domainLeader"

          Invoke-Expression $command
        }


        $teachers | ForEach-Object {

          $t = $_

          if(![string]::IsNullOrWhiteSpace($t)) {

            $command = "gam course $class add teacher $t"
            Invoke-Expression $command

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
        
        Invoke-Expression $command
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
    #$progressBarColor = $arg[3]

    #$Host.PrivateData.ProgressBackgroundColor=$progressBarColor

    Write-Progress -Activity $progressBarMessage -Status "Progress:" -PercentComplete ($progressCounter / $totalCount * 100)
  }

  Main

} 

Clear-ScriptPSSession
