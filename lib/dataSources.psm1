#$script:config = Get-Content -Raw -Path .\config.json | ConvertFrom-Json


function Get-ClassNames () {
  $classNamesCsv = Import-Csv -Path '.\tt\Class Names.csv'
  $studentLessonsCsv = Import-Csv -Path '.\tt\Student Lessons.csv'

  $timetableData = @{
    ttClasses = $classNamesCsv;
    ttStudentLessons = $studentLessonsCsv
  }

  $o = New-Object -Type PSObject -Property $timetableData

  $o.ttClasses
}