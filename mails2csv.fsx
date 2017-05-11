#r "packages/HtmlAgilityPack/lib/Net40/HtmlAgilityPack.dll"
#r "packages/MimeKitLite/lib/net40/MimeKitLite.dll"
#r "packages/CsvHelper/lib/net40/CsvHelper.dll"

open System.IO
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
  use fileWriter = new StreamWriter("data.csv")
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

(* Load all message, write single CSV file *)
let main (sourceDir: string) (dstFile: string) =
  Directory.GetFiles(sourceDir)
  |> Array.toList
  |> List.map loadHtmlFromMessage
  |> List.map extractTableData
  |> writeCsv dstFile

main "mails" "data.csv"
