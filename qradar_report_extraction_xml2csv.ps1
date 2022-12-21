$output=@()

Get-childitem *.xml | %{$file = $_.Name
  $author = (Select-Xml $file -XPath "//void[@property=""author""]").Node.string
  $CreationTime = (Select-Xml $file -XPath "//void[@property=""creationTime""]").Node.long
  $description = (Select-Xml $file -XPath "//void[@property=""descrition""]").Node.string
  $mailAdresse = (Select-Xml $file -XPath "//void[@property=""mailAddress""]").Node.string
  $name_id= (Select-Xml $file -XPath "//void[@property=""descrition""]").Node[-1].string
  $owner =  (Select-Xml $file -XPath "//void[@property=""owner""]").Node.string
  $runManually = (Select-Xml $file -XPath "//void[@property=""runManually""]").Node.boolean
  $scheduled = (Select-Xml $file -XPath "//void[@property=""scheduled""]").Node.boolean
  $title = (Select-Xml $file -XPath "//void[@property=""title""]").Node[-1].string

  $output +=($true | Select  @{N='file';E={$file}},
                  @{N='author';E={$author}},
                  @{N='CreationTime_epoch';E={$CreationTime}},
                  @{N='CreationTime';E={(Get-Date -Date "01-01-1970") + ([System.TimeSpan]::FromMilliseconds(($CreationTime)))}},
                  @{N='description';E={$description}},
                  @{N='mailAdresse';E={$mailAdresse}},
                  @{N='name_id';E={$name_id}},
                  @{N='owner';E={$owner}},
                  @{N='runManually';E={$runManually}},
                  @{N='scheduled';E={$scheduled}},
                  @{N='title';E={$title}}
          )
}
$output | Export-Csv -NoTypeInformation qradar_report_extraction.csv -Encoding UTF8
$output
