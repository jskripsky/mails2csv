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
    (texts.[0], texts.[1])

  let data =
    doc.DocumentNode.SelectNodes("//table[1]//tr")
    |> cast
    |> Seq.map processRow
    |> Seq.toList

  data

(* CSV output *)
let writeCsv (filename: string) (data: (string * string) list list) =
  let encoding = Encoding.GetEncoding(1252) // Windows-1252 for Excel CSV import
  use fileWriter = new StreamWriter(filename, false, encoding)
  use writer = new CsvWriter (fileWriter)

  let writeRow list =
    list
    |> List.map snd
    |> List.map (fun (s: string) -> writer.WriteField (s))
    |> ignore
    writer.NextRecord ()

  data
  |> List.map writeRow
  |> ignore
  
  fileWriter.Close ()

(* Load all messages, write CSV file *)
let mails2csv (sourceDir: string) (dstFile: string) =
  Directory.GetFiles(sourceDir)
  |> Array.toList
  |> List.map loadHtmlFromMessage
  |> List.map extractTableData
  |> writeCsv dstFile

let (<+>) x y = Path.Combine(x, y)

let getBaseDir () =
  let args = Environment.GetCommandLineArgs()
  Path.GetDirectoryName (args.[0])


[<EntryPoint>]
let main args =
  let baseDir = getBaseDir()
  let inDir = baseDir <+> "input"
  let outFile = baseDir <+> "output" <+> "data.csv"
  Console.WriteLine("Reading mails from '{0}', writing CSV file '{1}'...", inDir, outFile)
  mails2csv inDir outFile
  Console.WriteLine("Done. Please press enter to exit program.")
  Console.ReadLine() |> ignore
  0
