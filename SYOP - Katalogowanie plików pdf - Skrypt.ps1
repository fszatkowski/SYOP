#skrypt wykorzystuję bibilotekę itextsharp do wyciągnania informacji z plików pdf oraz aspose do ich renderowania

$script_path = Split-Path -Path $script:myinvocation.MyCommand.Definition -Parent
$aspose_name = "aspose.pdf.dll"
$aspose_path = Join-Path $script_path $aspose_name
$its_name = "itextsharp.dll"
$its_path = Join-Path $script_path $its_name

Add-Type -Path $aspose_path
Add-Type -Path $its_name

[System.Reflection.Assembly]::LoadFrom($aspose_path)

$work_dir = Get-Location
$work_dir.ToString()

# znajdź wszystkie pliki z rozszerzeniem pdf w folerze i podfolderach
$files = Get-ChildItem -Recurse | Where {$_.Extension -eq ".pdf"}

# sprawdź, ile plików znaleziono i wyświetl to
$number_of_files = $files.count
Write-Host "Starting to search pdf files"
Write-Host "Found $number_of_files files"

# zakończ działanie skryptu jeśli nie znaleziono żadnego pdfa
if (!$number_of_files)
{
    Write-Host "No document generated"
    Exit
}

Write-Host "Creating catalogue ..."

# utwórz dokument programu Word
$word = New-Object -ComObject "Word.Application"
$word.Visible = $False
$doc = $word.Documents.Add()

# do zapisywania tekstu potrzebny jest obiekt typu selection
$selection = $word.Selection

# za pomocą selection można też zmieniać czcionkę, rozmiar i kolor tekstu itp
$selection.WholeStory()
$selection.Font.Size = 12
$selection.Font.Name = "Times New Roman"
# enumeratory do różnych kolorów - https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb237558(v=office.12)
$selection.Font.Color = '0'

# przy formatowaniu tekstu użyty będzie tabulator i znak nowej linii
$tab = [char]9
$newline = [char]10


# stwórz pierwszą stronę - tytuł i spis treści
$selection.Font.Bold = 1
$selection.TypeText($tab + "Katalog plików pdf w " + $work_dir + $newline) 
$st = "$tab" +  "Spis Treści: " + "$newline"
$selection.TypeText($st)
$selection.Font.Bold = 0

$progress = [Decimal]0
$progress_quant = [Decimal]1/[Decimal]($number_of_files + 1)
$progress_value = [Decimal]0

$count = 1 
foreach ($file in $files)
{
    # utwórz nagłówek z nazwą pliku, ścieżką z hiperlinkiem do pliku oraz jego obrazem
    # wstaw tytuł oraz numer, formatuj do lewej strony
    $selection.ParagraphFormat.Alignment = 0
    $name = "$tab$count. " + ($file).Name + "$newline"
    $selection.TypeText($name)

    # ikrementuj licznik plików
    $count = $count + 1
}

# wyświetl progres :)
$progress = $progress + $progress_quant
if($progress -gt $progress_value * [Decimal]0.1)
{
    $progress_value = $progress_value + 1
    Write-Host -nonewline " | "
}

# przejdź na nową stronę
$selection.InsertBreak(2)

# dla każdego pliku przypisz numer i utwórz stonę w nowym dokumencie
$count = 1
foreach ($file in $files)
{
    # wyświetl progres :)
    $progress = $progress + $progress_quant
    if($progress -gt $progress_value * [Decimal]0.1)
    {
        $progress_value = $progress_value + 1
        Write-Host -nonewline " | "
    }
    # zapisz ścieżkę do pliku pdf
    [string]$pdfpath = ($file).fullname

    # utwórz nagłówek z nazwą pliku, ścieżką z hiperlinkiem do pliku oraz jego obrazem
    # wstaw tytuł oraz numer, formatuj do lewej strony
    $selection.ParagraphFormat.Alignment = 0
    $name = "$tab$count."
    $selection.TypeText($name)
    
    #wstaw hiperlink i ustaw jego właściwości    
    $hlink = $selection.Hyperlinks.Add($selection.Range, ($file).fullname)
    $hlink.Range.Font.Italic = $true
    $hlink.Range.Font.Name = 'Times New Roman'
    $hlink.Range.Font.Size = 12
    $hlink.Range.Font.Color = '10040115'

    # wstaw nową linię 
    $selection.TypeParagraph()

    #utwórz obiekt reader do wyciągnięcia z pdfa autora oraz tytułu
    $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $file.fullname
    $author = $reader.info["Author"]
    $title = $reader.info["Title"]
    $selection.TypeText("$tab Autor: $author, tytuł odczytany z pliku pdf: " + $reader.info["Title"])

    # wstaw nową linię 
    $selection.TypeParagraph()

    #formatuj do środka
    $selection.ParagraphFormat.Alignment = 1

    [string]$imgname = ($file).fullname + "_img.png"
    $paramOutput = " -sOutputFile=$imgname"

    # przekonwertuj pierwszą stronę na png za pomocą aspose.pdf

    $opts = New-Object Aspose.Pdf.RenderingOptions
    $opts.UseFontHinting = $true
    $pdfdoc = New-Object Aspose.Pdf.Document -ArgumentList $file.FullName
    $resolution = New-Object Aspose.Pdf.Devices.Resolution -ArgumentList 300
    $pngdev = New-Object Aspose.Pdf.Devices.PngDevice -ArgumentList $resolution
    $pngdev.RenderingOptions = $opts
    $page = $pdfdoc.Pages[1]
    $pngdev.Process($page, $imgname)

    # wstaw obraz do dokumentu
    $picture = $selection.InlineShapes.AddPicture($imgname)

    # usuń obraz, żeby nie śmiecić na komputerze użytkownika
    Remove-Item $imgname

    # przejdź na nową stronę, jeśli nie jesteśmy na ostatnim pliku
    if(!($count -eq $number_of_files))
    {
        $selection.InsertBreak(2)
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdbreaktype - różne typy łamania strony
    }

    # ikrementuj licznik plików
    $count = $count + 1
}

Write-Host " | finished"

# pobierz ścieżkę do folderu, w którym znajduje się skrypt skrypt i zapisz pod nią plik jako $output_name.docx
# zapytaj o nazwę pliku
$file_name = Read-Host -Prompt "Please enter catalouge name"
$outputPath = Join-Path $work_dir $file_name
$doc.SaveAs([string]$outputPath)

# wyświetl wiadomość o pomyślnym stworzeniu pliku
Write-Host "Catalogue created succesfully"

# zamknij dokument
$doc.Close()
$word.Quit()