$word = New-Object -ComObject Word.Application
$document = $word.Documents.Open("$(pwd)/docs/template.docx")
$selection = $word.Selection
$selection.TypeText("Hello, World!")
$document.SaveAs("$(pwd)/docs/hello-world.docx")
$document.Close()
$word.Quit()
