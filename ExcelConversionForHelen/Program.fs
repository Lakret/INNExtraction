#if INTERACTIVE
#r "stdole.dll"
#r "Microsoft.Office.Interop.Word"
#r "Microsoft.Office.Interop.Excel"
#r "System.Windows.Forms"
#endif

open System
open System.IO
open System.Collections.Generic
open System.Collections.Concurrent
open Microsoft.Office.Interop.Word
open Microsoft.Office.Interop.Excel
open System.Linq
open System.Text.RegularExpressions

let comarg x = ref (box x)
let baseDir = @"C:\Users\Lakret\Documents\Visual Studio 2012\Projects\ExcelConversionForHelen\ExcelConversionForHelen\скачанные лицензии"

let files = Directory.GetFiles(baseDir, "*.*", SearchOption.AllDirectories)
let extensions =
    set [|
        for file in files -> (Path.GetExtension file).ToLower()
    |]

//make subfolders for each type and move files with corresponding extensions to them
for extension in extensions do
    printfn "Now processing all %s" extension
    let selectedFiles = Directory.GetFiles(baseDir, "*" + extension, SearchOption.AllDirectories)
    printfn "%i files found..." selectedFiles.Length
    let newDir = sprintf @"%s\%s" baseDir extension
    Directory.CreateDirectory newDir |> ignore
    Array.iter
        (fun file -> 
            let newPath = sprintf @"%s\%s%s" newDir (Path.GetFileNameWithoutExtension file) extension
            printfn "path: %s" newPath
            File.Move(file, newPath))
        selectedFiles
        
let paths = new Dictionary<_, _>()
for folder in Directory.GetDirectories(baseDir) do
    let dirName = folder.Split([|'\\'|], StringSplitOptions.RemoveEmptyEntries).Last()
    if dirName.[0] = '.' then
        paths.Add(dirName, Directory.GetFiles folder)

let randGen = new Random()
let makeName() = 
    let arr = [| for _ in 1..10 -> randGen.Next(10).ToString() |] |> Array.map char 
    new String(arr) 

//convert XSL and XSLX to CSV
let excelDir = baseDir + @"\.xls"
let saveDir = baseDir + @"\.xls\processed"
let app = new ApplicationClass()
for excelFile in paths.[".xls"] do
    let makePath newDir = sprintf @"%s\%s" newDir <| makeName()
    printfn "Opening %s" excelFile
    app.Workbooks.Open excelFile |> ignore
    if app.Worksheets.Count > 1 then 
        printfn "Several worksheets book"
    for worksheetObj in app.Worksheets do
        let worksheet = worksheetObj :?> Worksheet
        let newPath = makePath saveDir
        printfn "Saving to %s" newPath
        worksheet.SaveAs(newPath, XlFileFormat.xlCSV)
        printfn "Saved"
app.Quit()


let isINN (str : string) =
    let rec normalize x = if x > 9 then normalize (x % 10) else x
    let getChecksum weights digits =
        let x = Array.map2 (*) weights digits |> Array.sum
        normalize <| x % 11
    let digits = Array.map (fun x -> x.ToString() |> Int32.Parse) <| str.ToCharArray()
    if digits.Length = 10 then
        getChecksum [| 2; 4; 10; 3; 5; 9; 4; 6; 8; 0 |] digits = digits.[9]
    elif digits.Length = 12 then
        let weights1 = [| 7; 2; 4; 10; 3; 5; 9; 4; 6; 8; 0 |]
        let weights2 = [| 3; 7; 2; 4; 10; 3; 5; 9; 4; 6; 8; 0 |]
        (getChecksum weights1 digits.[..10] = digits.[10]) && (getChecksum weights2 digits = digits.[11])
    else false

//extract all INN's from csv's
let extractAll path =
    let bag = new ConcurrentBag<_>()
    let regex = new Regex(@"\d{10,12}")
    Array.Parallel.iter 
        (fun csvFile ->
            let content = File.ReadAllText csvFile
            let matches = regex.Matches content
            [| for m in matches -> m.Value |] |> Array.filter isINN |> Array.iter bag.Add)
        (Directory.GetFiles(path, "*.csv"))
    bag

let res = extractAll (baseDir + @"\CSV")
let resSet = set [ yield! res ]     
File.WriteAllLines(baseDir + @"\fromExcel.txt", resSet.ToArray())

//converting rtf to txt
let rtfDir = baseDir + @"\.rtf"
let txtDir = baseDir + @"\txt"
Array.Parallel.iter
    (fun file -> 
        printfn "Processing %s" file
        let rtfBox = new System.Windows.Forms.RichTextBox()
        rtfBox.Rtf <- File.ReadAllText file
        File.WriteAllText(txtDir + @"\" + Path.GetFileNameWithoutExtension file + ".txt", rtfBox.Text))
    (Directory.GetFiles rtfDir)

//9909335836
//57791
let res' = extractAll txtDir
let resSet' = set [ yield! res' ]  
File.AppendAllLines(baseDir + @"\fromExcel.txt", resSet.ToArray())

//[<EntryPoint>]
//let main argv = 
//    printfn "%A" argv
//    0 // return an integer exit code
