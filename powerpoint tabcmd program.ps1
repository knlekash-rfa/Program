 # Generates a powerpoint file from a list of tableau urls.
 # Huge thanks to original author, Derrick Austin, you can find his original work @ https://www.interworks.com/blog/daustin>
 #>


$DebugPreference = "Continue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"
$ErrorActionPreference = "Continue"
$timeout = 300

# Base path of your script. Edit with your file path
$base ="C:\Users\User1\Client\3.PPT_QUERY"
$work_dir      = $base + '\temp'
Set-Location $work_dir

$download = 1
if ($download) 

# Enter your tableau online credentials below
{
    tabcmd login -s https://10ay.online.tableau.com/#/site/yoursite/home -u yourlogin@email.com -p password1
    if ($LastExitCode -ne 0) {
    echo "Login failed.  Goodbye.";
    break;
}

}
# File name to create temporary images, no edit necessary 

$file_name = "image"
# Import in your csv with your files, this requires your csv to be in the same base directory and titled 'Import'
$ps_demo = $base + "\Import.csv"    
$csv = Import-Csv -Path $ps_demo -Delimiter ","

$reports = $csv
$pdf_names = ""
# loop allows you to iterate through all of the tableau links, no edit necessary
for ($i = 0; $i -lt $reports.Count; $i++) 
{
    $current_report = $reports[$i]
    $png_name = $file_name + $i + ".png"
    Write-Host Exporting $png_name
    if ($download) 
    {  
        tabcmd export $current_report.url --png -f $png_name timeout $timeout
    }
# loop allows you to iterate through the coresponding imagemagick operations, no edit necessary

    $operations = $current_report.operations -split ","
    for ($j = 0; $j -lt $operations.Count; $j++) 
    {   name
        $cmd_magick="magick $($operations[$j])"
        $cmd_magick = $cmd_magick.replace("`$ORIG_NAME", $png_name)
        Write-Host Running: $cmd_magick
        Invoke-Expression $cmd_magick  
    }

    
}






#Sets the location for your input powerpoint with custom slide master and your output


$work_dir      = $base + '\temp\';
$input_pres    = $base + '\PPT_MASTER.pptx';
$output_folder = $base + '\output\';







$PPT         = New-Object -ComObject powerpoint.application;
$ori         = [Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal;
$my_ppt      = $PPT.Presentations.Open($input_pres, $false, $false, $false);
$num_slides  = $my_ppt.Slides.Count - 1;
$cover       = $my_ppt.Slides.Item(1);
$base_slide  = $my_ppt.Slides.Item(2);
$template    = $base_slide.CustomLayout;
$datestr     = Get-Date -f "DD MMMM yyyy";
$base_slide.delete();

# Generates the RGB color for the font you want to use.
$rgb = [long] (84 + (84 * 256) + (84 * 65536));



for ($i = 0; $i -lt $reports.Count; $i++){

    $filename  = $work_dir + $file_name + $i + ".png";
    
    
    $num_slides = $num_slides + 1;
    $new_slide  = $my_ppt.Slides.AddSlide($num_slides, $template);
    $new_pic    = $new_slide.shapes.AddPicture($filename, $false, $true, 0, 132);
    $new_pic.Title = 'Optional Title Here';
    $new_slide.Shapes.title.TextFrame.TextRange.Text = 'Optional Title Here';
    
    $new_pic.LockAspectRatio = $false;
    $new_pic.Width  = $new_slide.CustomLayout.Width + 1;
    $new_pic.Height = 350;
    }

#Sets the name of your newly created powerpoint
$name = $output_folder + "Monthly Powerpoint (+$datestr)";
$my_ppt.SaveCopyAs($name);
$my_ppt.Close();