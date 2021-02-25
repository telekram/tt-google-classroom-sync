param(
  [string]$CsvPath=$null,
  [string]$ShowSubjects=$null,
  [string]$AddSubjects=$null,
  [string]$AddClasses=$null,
  [string]$AddTeachersToSubjects=$null,
  [string]$AddTeachersToClasses=$null,
  [string]$AddStudentsToClasses=$null,
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

  $GAM = [PSCustomObject]@{}
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

    if(!$null -eq $addStudentsToClasses){
      Add-StudentsToClasses($addStudentsToClasses)
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
    

      $subjectCode = $_.SubjectCode
      $subjectName = $_.SubjectName


      $progressCounter = $progressCounter + 1
      $progressBarMessage = "Adding course: $subjectCode - $subjectName"


      Get-ProgressBar (
        $progressCounter,
        @($s).Count,
        $progressBarMessage,
        'Magenta',
        0
      )

      $GAM.PublishCourse($course)
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

    $subjectCount = @($s).Count

    $s | ForEach-Object {

      $progressCounter1 = 0
      $classCodesCount = $_.ClassCodes.Count
      
      $subjectCode = $_.subjectCode
      $subjectName = $_.SubjectName
      $faculty = $_.FacultyName

      
      $progressBarMessage = "Subject: $subjectCode - $subjectName"
      
      Get-ProgressBar (
        $progressCounter0,
        $subjectCount,
        $progressBarMessage,
        'Magenta',
        0
      )

      $progressCounter0 = $progressCounter0 + 1

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
        
        $GAM.PublishCourse($course)
      }
    }
  }


  function Add-TeachersToSubjects($subject) {

    $s = $DataSet.Subjects | 
    Where-Object { $_.SubjectCode -like "*$subject*" } 

    if(!$s) {
      Write-Host "Subject(s): '$subject' not found"
      exit
    }

    $progressCounter0 = 0
    $subjectCount = @($s).Count

    $s | ForEach-Object {

      $subjectCode = $_.SubjectCode
      $teachers = $_.Teachers
      $domainLeader = $_.DomainLeader

      $progressBarMessage0 = "Subject: $subjectCode"
      $progressCounter0 = $progressCounter0 + 1

      if(![string]::IsNullOrWhiteSpace($domainLeader)) {

        if($teachers -notcontains $domainLeader){
          $teachers += $domainLeader
        }
      }

      Get-ProgressBar (
        $progressCounter0,
        $subjectCount,
        $progressBarMessage0,
        'Magenta',
        0
      )

      $progressCounter1 = 0
      $teacherCount = $teachers.Count

      $teachers | ForEach-Object {
        
        $teacher = $_

        $courseParticipant = [PSCustomObject]@{
          Course = $subjectCode
          Type = 'Teacher'
          Participant = $teacher
        }

        $progressCounter1 = $progressCounter1 + 1
        $progressBarMessage1 = "Adding $teacher to course: " + $academicYear + '-' + $subjectCode

        Get-ProgressBar (
          $progressCounter1,
          $teacherCount,
          $progressBarMessage1,
          'Magenta',
          1
        )
  
        $GAM.AddCourseParticipant($courseParticipant)
      }  
    }
  }

  function Add-TeachersToClasses($subject) {

    $s = $DataSet.Subjects | 
      Where-Object { $_.SubjectCode -like "*$subject*" } 

    if(!$s) {
      Write-Host "Subject(s): '$subject' not found"
      exit
    }

    $progressCounter0 = 0
    $subjectCount = @($s).Count

    $s | ForEach-Object {

      $subjectCode = $_.SubjectCode
      $domainLeader = $_.DomainLeader
      $teachers = $_.Teachers
      $classCodes = $_.ClassCodes

      $progressBarMessage0 = "Working on subject: $subjectCode"
      $progressCounter0 = $progressCounter0 + 1

      Get-ProgressBar (
        $progressCounter0,
        $subjectCount,
        $progressBarMessage0,
        'Magenta',
        0
      )

      $progressCounter1 = 0
      $classCount = @($classCodes).Count

      $classCodes | ForEach-Object {
        
        $class = $academicYear + '-' + $_

        if(![string]::IsNullOrWhiteSpace($domainLeader)) {

          $courseParticipant = [PSCustomObject]@{
            Course = $class
            Type = 'Teacher'
            Participant = $domainLeader
          }

          $GAM.AddCourseParticipant($courseParticipant)
        }

        $progressCounter1 = $progressCounter1 + 1
        $progressBarMessage1 = "Working on class: $class"

        Get-ProgressBar (
          $progressCounter1,
          $classCount,
          $progressBarMessage1,
          'Magenta',
          1
        )

        $progressCounter2 = 0
        $teacherCount = @($teachers).Count

        $teachers | ForEach-Object {

          $teacher = $_

          $progressCounter2 = $progressCounter2 + 1
          $progressBarMessage2 = "Adding $teacher to Google Course $class"

          Get-ProgressBar (
            $progressCounter2,
            $teacherCount,
            $progressBarMessage2,
            'Magenta',
            2
          )


          if(![string]::IsNullOrWhiteSpace($teacher)) {

            $courseParticipant = [PSCustomObject]@{
              Course = $class
              Type = 'Teacher'
              Participant = $teacher
            }
  
            $GAM.AddCourseParticipant($courseParticipant)
          }
        }
      }  
    }
  }


  function Add-StudentsToClasses($class) {

    $c = $DataSet.Classes | 
    Where-Object { $_.ClassCode -like "*$class*" } 

    if(!$c) {
      Write-Host "Subject(s): '$class' not found"
      exit
    }
    
    $progressCounter0 = 0

    $c | ForEach-Object {

      $class = $academicYear + '-' + $_.ClassCode
     
      $students = $_.StudentCodes

      $progressCounter0 = $progressCounter0 + 1
      $progressBarMessage0 = "Google Course: $class"

      Get-ProgressBar (
        $progressCounter0,
        $DataSet.Classes.Count,
        $progressBarMessage0,
        'Magenta',
        0
      )

      $progressCounter1 = 0
      $studentCount = $students.Count

      $students | ForEach-Object {

        $student = $_

        $progressCounter1 = $progressCounter1 + 1
        $progressBarMessage1 = "Adding student $student"


        Get-ProgressBar (
          $progressCounter1,
          $studentCount,
          $progressBarMessage1,
          'Magenta',
          1
        )

        $courseParticipant = [PSCustomObject]@{
          Course = $class
          Type = 'Student'
          Participant = $student
        }

        $GAM.AddCourseParticipant($courseParticipant)
      }
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



  $publishCourse = {

    param([PSCustomObject]$courseAttributes)

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
      Write-Host $cmd
      Start-Sleep -Seconds 1
    }
  }

  $GAM | Add-Member -MemberType ScriptMethod -Name PublishCourse -Value $publishCourse



  $addCourseParticipant = {

    param([PSCustomObject]$courseParticipant)

    if([string]::IsNullOrWhiteSpace($courseParticipant.Type)){
      Write-Host "Add course participant error: Participant type not defined (stduent/teacher). Script will exit"
      exit
    }

    if([string]::IsNullOrWhiteSpace($courseParticipant.Course)){
      Write-Host "Add course participant error: Course not specified. Script will exit"
      exit
    }

    $course = $courseParticipant.Course
    $type = $courseParticipant.Type
    $participant = $courseParticipant.Participant

    $cmd = $null

    if($courseParticipant.Type -eq 'Student') {

      $cmd = "gam course $course add $type $participant"

    } elseif ($courseParticipant.Type -eq 'Teacher') {

      $cmd = "gam course $course add $type $participant"

    }

    if(!$isSimulatingCommands) {
      Invoke-Expression $cmd
    } else {
      Write-Host $cmd
      Start-Sleep -Seconds 1
    }
  }

  $GAM | Add-Member -MemberType ScriptMethod -Name AddCourseParticipant -Value $addCourseParticipant


  function Get-ProgressBar ($arg) {    
    
    $progressCounter = $arg[0]
    $totalCount = $arg[1]
    $progressBarMessage = $arg[2]
    $progressBarColor = $arg[3]
    $progressBarId = $arg[4]

    #$Host.PrivateData.ProgressBackgroundColor=$progressBarColor

    switch ($progressBarId) {
      0 {
        Write-Progress -Id $progressBarId -Activity $progressBarMessage -Status "Progress:" -PercentComplete ($progressCounter / $totalCount * 100)
      }
      1 {
        Write-Progress -Id $progressBarId -ParentId 0 -Activity $progressBarMessage -Status "Progress:" -PercentComplete ($progressCounter / $totalCount * 100)
      }
      2 {
        Write-Progress -Id $progressBarId -ParentId 1 -Activity $progressBarMessage -Status "Progress:" -PercentComplete ($progressCounter / $totalCount * 100)
      }
    }  
  }

  Main

} 

Clear-ScriptPSSession
