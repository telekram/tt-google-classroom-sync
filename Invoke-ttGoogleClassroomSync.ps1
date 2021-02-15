param(
  [string]$CsvPath=$null,
  [string]$ShowSubjects=$null,
  [string]$AddSubjects=$null,
  [string]$AddClasses=$null,
  [string]$AddTeachersToSubjects=$null,
  [string]$AddTeachersToClasses=$null,
  [switch]$AddStudentsToClasses,
  [switch]$AddCompositeClasses,
  [switch]$GetRemoteCourses,
  [switch]$TestGamCommand,
  [switch]$SimulateCommands
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
  $testGamCommand = $Using:TestGamCommand
  $isSimulatingCommands = $Using:SimulateCommands

  
  function Main {

    if($testGamCommand){
      Test-GamCommand
    }

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

    if(!$null -eq $addTeachersToClasses) {
      Add-TeachersToClasses($addTeachersToClasses)
    }

    if($addStudentsToClasses){
      Add-StudentsToClasses
    }

    if($addCompositeClasses){
      Add-CompositeClasses
    }

  }

  function Test-GamCommand {
    #gam info course '2021-1ACC115'
    $p = gam print course-participants course '2021-1ACC115' | Out-String
    
    $part = $p | ConvertFrom-Csv -Delim ','
    
    #$part
    $part | ForEach-Object {
      $_.'profile.emailAddress'
      
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
      Write-Host 'e' -NoNewline
    }

    $script:CloudCourses
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
        'Magenta',
        0
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
      
    } elseif ($c.Type -eq 'Class') {

      $alias = $academicYear + '-' + $c.Code
      $name = $c.Code
    }
    
    $section = $c.Name -Replace '[^a-zA-Z0-9-_ ]', ''
    $description = 'Subject Domain: ' + $c.Faculty + ' - ' + $section
    $room = $academicYear

    $cmd = "gam create course alias $alias name '$name' section '$section' description '$description' heading $alias room $room teacher $classroomAdmin status active"

    if(!$isSimulatingCommands) {
      Invoke-Expression $cmd
    } else {
      #Write-Host $cmd
    }
  }


  function Add-ClassCoursesToGoogle($subject) {

    $progressCounter0 = 0
    
    $s = $DataSet.Subjects | 
      Where-Object { $_.SubjectCode -like "*$subject*" } 

    if(!$s) {
      Write-Host "Subject(s): '$subject' not found"
      exit
    }

    $s | ForEach-Object {

      $progressCounter1 = 0
      $classCodesCount = $_.ClassCodes.Count
      
      $subjectName = $_.SubjectName
      $faculty = $_.FacultyName

      $progressBarMessage = "Subject: $subjectName"

      Get-ProgressBar (
        $progressCounter0,
        $DataSet.Subjects.Count,
        $progressBarMessage,
        'Magenta',
        0
      )

      $_.ClassCodes | ForEach-Object {

        $cc = $_

        $course = [PSCustomObject]@{
          Type = 'Class'
          Code = $cc
          Name = $subjectName
          Faculty = $faculty
        }

        $progressBarMessage = "Publishing course: $cc"

        $progressCounter1 = $progressCounter1 + 1

        Get-ProgressBar (
          $progressCounter1,
          $classCodesCount,
          $progressBarMessage,
          'Magenta',
          1
        )
        
        Publish-Course($course)   
      }
      
      $progressCounter0 = $progressCounter0 + 1
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

        if(!$isSimulatingCommands) {
          Write-Host $cmd
          $cmd = "gam course $classAlias add teacher $teacher"
          Invoke-Expression $cmd
        } else {
          Write-Host $cmd
        }
      }


      $_.ClassCodes | ForEach-Object {
        
        $classCode = $_
        
        $DataSet.Classes.$classCode | ForEach-Object {

          $student = $_

          $command = "gam course $classAlias add student $student"
          #Invoke-Expression $command

          Write-Host "Theres a bug here command is disabled. Investigate"
          Write-Host $command 
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

  function Add-TeachersToSubjects($subject) {

    $progressCounter = 0
    
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
      
      $progressBarMessage = $command
      $progressCounter = $progressCounter + 1

      Get-ProgressBar (
        $progressCounter,
        $DataSet.Subjects.Count,
        $progressBarMessage,
        'Magenta'
      )
    }
  }

  function Add-TeachersToClasses($subject) {

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

        $progressBarMessage = $command

        Get-ProgressBar (
          $progressCounter,
          $DataSet.Subjects.Count,
          $progressBarMessage,
          'Magenta'
        )
      }  

      
      $progressCounter = $progressCounter + 1

      # Get-ProgressBar (
      #   $progressCounter,
      #   $DataSet.Subjects.Count,
      #   $progressBarMessage,
      #   'Magenta'
      # )
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

        $progressBarMessage = $command

        Invoke-Expression $command

        Get-ProgressBar (
          $progressCounter,
          $DataSet.Classes.Count,
          $progressBarMessage,
          'Magenta'
        )
      }

      
      $progressCounter = $progressCounter + 1

      # Get-ProgressBar (
      #   $progressCounter,
      #   $DataSet.Classes.Count,
      #   $progressBarMessage,
      #   'Magenta'
      # )
    }
  }

  function Get-ProgressBar ($arg) {    
    
    $progressCounter = $arg[0]
    $totalCount = $arg[1]
    $progressBarMessage = $arg[2]
    $progressBarColor = $arg[3]
    $progressBarId = $arg[4]

    #$Host.PrivateData.ProgressBackgroundColor=$progressBarColor

    if ($progressBarId -gt 0) {
      Write-Progress -Id $progressBarId -ParentId 0 -Activity $progressBarMessage -Status "Progress:" -PercentComplete ($progressCounter / $totalCount * 100)
    } else {
      Write-Progress -Id $progressBarId -Activity $progressBarMessage -Status "Progress:" -PercentComplete ($progressCounter / $totalCount * 100)
    }

    
  }

  Main

} 

Clear-ScriptPSSession
