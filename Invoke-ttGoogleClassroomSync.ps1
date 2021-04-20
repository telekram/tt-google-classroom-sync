param(
  [string]$CsvPath=$null,
  [string]$ShowSubjects=$null,
  [string]$AddSubjects=$null,
  [string]$AddClasses=$null,
  [string]$AddTeachersToSubjects=$null,
  [string]$AddTeachersToClasses=$null,
  [string]$AddStudentsToClasses=$null,
  [string]$AddTeacherToClasses=$null,
  [string]$Teacher=$null,
  [switch]$AddCompositeClasses,
  [switch]$GetRemoteCourses,
  [string]$FindRemoteCourse=$null,
  [string]$ArchiveCourses=$null,
  [string]$TestGamCommand=$null,
  [switch]$SimulateCommands
) 


$ScriptParameters = [PSCustomObject]@{
  CsvPath = $CsvPath
  ShowSubjects = $ShowSubjects
  AddSubjects = $AddSubjects
  AddClasses = $AddClasses
  AddTeachersToSubjects = $AddTeachersToSubjects
  AddTeachersToClasses = $AddTeachersToClasses
  AddStudentsToClasses = $AddStudentsToClasses
  AddTeacherToClasses = $AddTeacherToClasses
  Teacher = $Teacher
  AddCompositeClasses = $AddCompositeClasses
  GetRemoteCourses = $GetRemoteCourses
  FindRemoteCourse = $FindRemoteCourse
  ArchiveCourses = $ArchiveCourses
  IsSimulateCommands = $SimulateCommands
  TestGamCommand = $TestGamCommand
}

$script:config = Get-Content -Raw -Path .\config.json | ConvertFrom-Json
Import-Module .\lib\dataSources.psm1 -Force -Scope Local
Import-Module .\lib\sessionManager.psm1 -Force -Scope Local

$DS = Get-DataSourceObject($scriptParameters.CsvPath)

$CA = $script:config.classroomAdmin
$AY = $script:config.academicYear
 
$session = Get-ScriptPSSession

Invoke-Command -Session $session -ScriptBlock {

  $GAM = [PSCustomObject]@{}
  $DataSet = $Using:DS

  $academicYear = $Using:AY
  $classroomAdmin = $Using:CA

  $scriptParameters = $Using:ScriptParameters 

  
  function Main {

    if(!$null -eq $scriptParameters.TestGamCommand){
      Test-GamCommand($scriptParameters.TestGamCommand)
    }

    if($scriptParameters.GetRemoteCourses){
      Get-CoursesFromGoogle
    }

    if(!$null -eq $scriptParameters.ShowSubjects) {
      Show-Subject($scriptParameters.ShowSubjects)
    } 

    if(!$null -eq $scriptParameters.AddSubjects) {
      Add-SujectCoursesToGoogle($scriptParameters.AddSubjects)
    } 

    if(!$null -eq $scriptParameters.AddClasses) {
      Add-ClassCoursesToGoogle($scriptParameters.AddClasses)
    }

    if(!$null -eq $scriptParameters.AddTeachersToSubjects) {
      Add-TeachersToSubjects($scriptParameters.AddTeachersToSubjects)
    }

    if(!$null -eq $scriptParameters.AddTeachersToClasses) {
      Add-TeachersToClasses($scriptParameters.AddTeachersToClasses)
    }

    if(!$null -eq $scriptParameters.AddTeacherToClasses) {
      if(!$null -eq $scriptParameters.Teacher) {
        Add-TeacherToClasses($scriptParameters.AddTeacherToClasses, $scriptParameters.Teacher)
      } else {
        Write-Warning "Teacher parameter required"
      }
    }

    if(!$null -eq $scriptParameters.AddStudentsToClasses){
      Add-StudentsToClasses($scriptParameters.AddStudentsToClasses)
    }

    if(!$null -eq $scriptParameters.ArchiveCourses) {
      Remove-Courses($scriptParameters.ArchiveCourses)
    }

    if(!$null -eq $scriptParameters.FindRemoteCourse){
      $GAM.FindRemoteCourse($scriptParameters.FindRemoteCourse)
    }

    if($scriptParameters.AddCompositeClasses){
      Add-CompositeClasses
    }
  }

  function Test-GamCommand($subject) {
    
  }


  function Get-CoursesFromGoogle {

    [System.Collections.ArrayList]$script:CloudCourses = @()

    $gCourses = gam print courses teacher $classroomAdmin 2> $null | Out-String
    $courses = $gCourses | ConvertFrom-Csv -Delim ','

    $progressCounter0 = $progressCounter0 + 1
    $totalCloudCourses = @($courses).Count

    $courses | ForEach-Object {

      $courseAlias = $_.DescriptionHeading

      $progressCounter0 = $progressCounter0 + 1
      $progressBarMessage0 = "Adding $courseAlias to CloudCourses object"

      Get-ProgressBar (
        $progressCounter0,
        $totalCloudCourses,
        $progressBarMessage0,
        0
      )

    
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
  }


  function Show-Subject($subject) {

    $s = $DataSet.Subjects | 
      Where-Object { $_.SubjectCode -like "*$subject*" -or $_.SubjectName -like "*$subject*" } 

    if(!$s) {
      Write-Host "Subject(s): '$subject' not found"
      exit
    }

    $s
  }

  function Add-SujectCoursesToGoogle($subject) {
    
    $subjects = $DataSet.Subjects |
      Where-Object { $_.SubjectCode -like "*$subject*" } 

    if(!$subjects) {
      Write-Host "Subject(s): '$subject' not found"
      exit
    }

    $progressCounter = 0
    $subjectCount = @($s).Count

    $subjects | ForEach-Object {

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
        $subjectCount,
        $progressBarMessage,
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
        0
      )

      $progressCounter0 = $progressCounter0 + 1

      $_.ClassCodes | ForEach-Object {

        $class = $_.Class

        $course = [PSCustomObject]@{
          Type = 'Class'
          Code = $class
          Name = $subjectName
          Faculty = $faculty
        }

        $progressBarMessage = "Publishing course: $class"

        $progressCounter1 = $progressCounter1 + 1

        Get-ProgressBar (
          $progressCounter1,
          $classCodesCount,
          $progressBarMessage,
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
      [array]$teachers = $_.Teachers
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
      [array]$teachers = $_.Teachers
      $classCodes = $_.ClassCodes

      if(![string]::IsNullOrWhiteSpace($domainLeader)) {

        if($teachers -notcontains $domainLeader){
          $teachers += $domainLeader
        }
      }

      $progressBarMessage0 = "Working on subject: $subjectCode"
      $progressCounter0 = $progressCounter0 + 1

      Get-ProgressBar (
        $progressCounter0,
        $subjectCount,
        $progressBarMessage0,
        0
      )

      $progressCounter1 = 0
      $classCount = @($classCodes).Count

      $classCodes | ForEach-Object {
        
        $class = $academicYear + '-' + $_.Class

        $progressCounter1 = $progressCounter1 + 1
        $progressBarMessage1 = "Working on class: $class"

        Get-ProgressBar (
          $progressCounter1,
          $classCount,
          $progressBarMessage1,
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

  function Add-TeacherToClasses($arg) {

    $class = $arg[0]
    $teacher = $arg[1]

    $subjects = $DataSet.Subjects | 
    Where-Object { $_.SubjectCode -like "*$class*" -or $_.SubjectName -like "*$class*" } 

    if(!$subjects) {
      Write-Host "Subject(s): '$class' not found"
      exit
    }

    $progressCounter0 = 0
    $classCodeCount = $subjects.ClassCodes.Count

    $subjects.ClassCodes | ForEach-Object {

      $class = $academicYear + '-' + $_.Class

      $progressCounter0 += 1
      $progressBarMessage0 = "Adding teacher $teacher to course $class"

      Get-ProgressBar (
        $progressCounter0,
        $classCodeCount,
        $progressBarMessage0,
        0
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


  function Add-StudentsToClasses($class) {

    $classCodes = $DataSet.Subjects.ClassCodes | 
    Where-Object { $_.Class -like "*$class*" } 

    if(!$classCodes) {
      Write-Host "Subject(s): '$class' not found"
      exit
    }
    
    $progressCounter0 = 0
    $classCodesTotal = @($classCodes).Count

    $classCodes | ForEach-Object {

      $class = $academicYear + '-' + $_.Class
     
      $students = $_.Students

      $progressCounter0 = $progressCounter0 + 1
      $progressBarMessage0 = "Google Course: $class"

      Get-ProgressBar (
        $progressCounter0,
        $classCodesTotal,
        $progressBarMessage0,
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

    $progressCounter0 = 0
    $compositeClassCount = @($DataSet.CompositeClasses).Count

    $DataSet.CompositeClasses | ForEach-Object {

      $compositeClassCodes = $_.ClassCodes

      $subjectCode = $_.SubjectCode

      $classAlias = "$academicYear-$subjectCode"

      $course = [PSCustomObject]@{
        Type = 'Class'
        Code = $subjectCode
        Name = $_.SubjectName
        Faculty = ''
      }

      $progressBarMessage = "Adding composite class course: $classAlias"
      $progressCounter0 = $progressCounter + 1

      Get-ProgressBar (
        $progressCounter0,
        $compositeClassCount,
        $progressBarMessage,
        0
      )

      $GAM.PublishCourse($course)

      $progressCounter1 = 0
      $teacherCount = @($_.Teachers).Count

      $_.Teachers | ForEach-Object {

        $teacher = $_

        $courseParticipant = [PSCustomObject]@{
          Course = $classAlias
          Type = 'Teacher'
          Participant = $teacher
        }

        $progressCounter1 = $progressCounter + 1
        $progressBarMessage = "Adding teacher $teacher to $classAlias"
        
        Get-ProgressBar (
          $progressCounter1,
          $teacherCount,
          $progressBarMessage,
          1
        )

        $GAM.AddCourseParticipant($courseParticipant)
      }

      $DataSet.Subjects.ClassCodes | ForEach-Object {

        if ($compositeClassCodes -contains $_.Class){


          $progressCounter1 = 0
          $studentCount = $_.Students.Count

          $_.Students | ForEach-Object {

            $student = $_

            $courseParticipant = [PSCustomObject]@{
              Course = $classAlias
              Type = 'Student'
              Participant = $student
            }

            $progressCounter1 = $progressCounter + 1
            $progressBarMessage = "Adding student $student to $classAlias"
      
            Get-ProgressBar (
              $progressCounter1,
              $studentCount,
              $progressBarMessage,
              1
            )

            $GAM.AddCourseParticipant($courseParticipant)
          }
        }
      }
    }
  }

  function Remove-Courses($course) {

    Write-Host "Fetching courses from Google"
    
    Get-CoursesFromGoogle
    
    $returnedCourses = $script:CloudCourses | Where-Object { 
      $_.DescriptionHeading -like "$academicYear-*$course*" -and $_.CourseState -eq 'ACTIVE'
    }

    $returnedCourses

    if($returnedCourses) {
      $answer = Read-Host -Prompt 'Do you wish to archive these course(s)? (y/n)'
    } else {
      Write-Host "No courses found for '$course'"
    }

    if($answer.ToLower() -eq 'y') {

      $progressCounter0 = 0
      $courseCount = @($returnedCourses).Count

      $returnedCourses | ForEach-Object {

        $courseAlias = $_.DescriptionHeading

        $progressCounter0 += 1
        $progressBarMessage = "Archiving $courseAlias" 

        Get-ProgressBar (
          $progressCounter0,
          $courseCount,
          $progressBarMessage,
          0
        )
        
        $GAM.ArchiveCourse($courseAlias)
      }
    }
  }


  $publishCourse = {

    param([PSCustomObject]$courseAttributes)

    $c = $courseAttributes

    if([string]::IsNullOrWhiteSpace($c.Type)){
      Write-Host "Publish course error: course type not defined. Script will exit"
      exit
    }

    switch ($c.Type) {

      'Subject' {
        $alias = $c.Code
        $name = $c.Code + ' (Teachers)'
      }

      'Class' {

        $alias = $academicYear + '-' + $c.Code
        $name = $c.Code
      }
    }

    
    $section = $c.Name -Replace '[^a-zA-Z0-9-_ ]', ''
    $description = 'Subject Domain: ' + $c.Faculty + ' - ' + $section
    $room = $academicYear

    $cmd = "gam create course alias $alias name '$name' section '$section' description '$description' heading $alias room $room teacher $classroomAdmin status active"

    if(!$scriptParameters.IsSimulateCommands) {
      Invoke-Expression $cmd
    } else {
      Write-Host $cmd
      Start-Sleep -Seconds 5
    }
  }

  $GAM | Add-Member -MemberType ScriptMethod -Name PublishCourse -Value $publishCourse

  $archiveCourse = {

    param([string]$courseAlias)

    $alias = $courseAlias

    if([string]::IsNullOrWhiteSpace($alias)){
      Write-Host "Delete course error: course type not defined. Script will exit"
      exit
    }

    $cmd = "gam update course $alias status archive"

    if(!$scriptParameters.IsSimulateCommands) {
      Invoke-Expression $cmd
    } else {
      Write-Host $cmd
      Start-Sleep -Seconds 5
    }
  }

  $GAM | Add-Member -MemberType ScriptMethod -Name ArchiveCourse -Value $archiveCourse

  $addCourseParticipant = {

    param([PSCustomObject]$courseParticipant)

    if([string]::IsNullOrWhiteSpace($courseParticipant.Type)){
      Write-Host "Add course participant error: Participant type not defined (stduent/teacher). Script will now exit"
      exit
    }

    if([string]::IsNullOrWhiteSpace($courseParticipant.Course)){
      Write-Host "Add course participant error: Course not specified. Script will now exit"
      exit
    }

    $course = $courseParticipant.Course
    $type = $courseParticipant.Type
    $participant = $courseParticipant.Participant

    $cmd = "gam course $course add $type $participant"

    if(!$scriptParameters.IsSimulateCommands) {
      Invoke-Expression $cmd
    } else {
      Write-Host $cmd
      Start-Sleep -Seconds 5
    }
  }

  $GAM | Add-Member -MemberType ScriptMethod -Name AddCourseParticipant -Value $addCourseParticipant

  $findRemoteCourse = {
    param([string]$subject)

    Get-CoursesFromGoogle
    
    $returnedSubjects = $script:CloudCourses | Where-Object { 
      $_.DescriptionHeading -like "*$subject*" 
    }

    $returnedSubjects
  }

  $GAM | Add-Member -MemberType ScriptMethod -Name FindRemoteCourse -Value $findRemoteCourse
  
  function Get-ProgressBar ($arg) {    
    
    $progressCounter = $arg[0]
    $totalCount = $arg[1]
    $progressBarMessage = $arg[2]
    $progressBarId = $arg[3]

    if ($progressBarMessage -eq 0) {
      Write-Progress -Id $progressBarId -Activity $progressBarMessage -Status "Progress:" -PercentComplete ($progressCounter / $totalCount * 100)
    } else {

      $parentId = $progressBarId - 1
      Write-Progress -Id $progressBarId -ParentId $parentId -Activity $progressBarMessage -Status "Progress:" -PercentComplete ($progressCounter / $totalCount * 100)
    }
  }

  Main

} 

Clear-ScriptPSSession
