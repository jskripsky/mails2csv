#if INTERACTIVE
#r "packages/HtmlAgilityPack/lib/Net40/HtmlAgilityPack.dll"
#r "packages/MimeKitLite/lib/net40/MimeKitLite.dll"
#r "packages/CsvHelper/lib/net40/CsvHelper.dll"
#endif

open System
open System.IO
open System.Text

open MimeKit
open HtmlAgilityPack
open CsvHelper

(* Mime/HTML input *)
let loadHtmlFromMessage (msgFilename: string) =
  let msg = MimeMessage.Load(msgFilename)
  msg.HtmlBody

(*HTML table parsing *)
let extractTableData (html: string) =
  let doc = new HtmlDocument ()
  doc.LoadHtml(html)

  let cast = Seq.cast<HtmlNode>
  let trim (s: string) = s.Replace("&nbsp;", "").Trim()

  let processRow (r: HtmlNode) =
    let texts =
      r.SelectNodes("td")
      |> cast
      |> Seq.map (fun n -> n.InnerText |> trim)
      |> Seq.toList
    texts.[1]

  let data =
    doc.DocumentNode.SelectNodes("//table[1]//tr")
    |> cast
    |> Seq.map processRow
    |> Seq.toList

  data

let (<+>) x y = Path.Combine(x, y)

let fixMissingRow (list: string list) =
  let insertAt (list: string list) pos item =
    List.append (list.[..pos - 1]) (item::list.[pos..])

  if list.Length = 6 then
    list
  else
    insertAt list 4 "-"

let processMessageFile (log: TextWriter) (file: string) =
  log.Write(String.Format("Reading file '{0}'... ", Path.GetFileName(file)))
  try
    let html = loadHtmlFromMessage file
    log.Write("extracting table data... ")
    let data = extractTableData html |> fixMissingRow
    log.WriteLine("done.")

    let dstPath = Path.GetDirectoryName(file) <+> "processed" <+> Path.GetFileName(file)
    log.WriteLine(String.Format("Moving {0} to {1}.", file, dstPath))
    File.Move(file, dstPath)

    Some data
  with
    | exc ->
      log.WriteLine(String.Format("Error: {0}", exc))
      log.WriteLine()
      None

(* CSV output *)
let writeCsv (filename: string) (data: string list list) =
  let encoding = Encoding.GetEncoding(1252) // Windows-1252 for Excel CSV import
  use fileWriter = new StreamWriter(filename, false, encoding)
  use writer = new CsvWriter (fileWriter)

  let writeRow list =
    list
    |> List.map (fun (s: string) -> writer.WriteField (s))
    |> ignore
    writer.NextRecord ()

  data
  |> List.map writeRow
  |> ignore
  
  fileWriter.Close ()
  

(* Load all messages, write CSV file *)
let mails2csv (log: TextWriter) (sourceDir: string) (dstFile: string) =
  let proc =processMessageFile log

  Directory.GetFiles(sourceDir)
  |> Array.toList
  |> List.choose proc
  |> writeCsv dstFile

let getBaseDir () =
  let args = Environment.GetCommandLineArgs()
  Path.GetDirectoryName (args.[0])


[<EntryPoint>]
let main args =
  let ts = DateTime.Now.ToString("yyyyMMdd-HHmmss")

  let baseDir = getBaseDir()
  let inDir = baseDir <+> "input"
  let outFile = baseDir <+> "output" <+> "data-" + ts + ".csv"
  let logFile = baseDir <+> "output" <+> "log-" + ts + ".txt"

  Console.WriteLine("Reading mails from '{0}.", inDir)
  Console.WriteLine("Writing to CSV file '{0}'.", outFile)
  Console.WriteLine("Writing to Log file '{0}'.", logFile)
  Console.WriteLine()


  use log = new StreamWriter(logFile)
  mails2csv log inDir outFile

  Console.WriteLine()
  Console.WriteLine("All done. Please press enter to exit program.")
  Console.ReadLine() |> ignore
  0
