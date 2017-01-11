#r "../packages/DocumentFormat.OpenXml/lib/DocumentFormat.OpenXml.dll"
#r "../packages/NeXL/lib/net45/NeXL.XlInterop.dll"
#r "../packages/NeXL/lib/net45/NeXL.ManagedXll.dll"
#r "../packages/FSharp.Data/lib/net40/FSharp.Data.dll"
#r "../packages/Newtonsoft.Json/lib/net45/Newtonsoft.Json.dll"
#r "../packages/Deedle/lib/net40/Deedle.dll"

#load "Types.fs"

open NeXL.ManagedXll
open NeXL.XlInterop
open System
open System.IO
open System.Runtime.InteropServices
open System.Data
open FSharp.Data
open FSharp.Data.JsonExtensions
open FSharp.Data.HtmlExtensions
open Newtonsoft.Json
open Newtonsoft.Json.Linq
open NeXL.Worldbank
open Deedle

let formatPrm = Some ("format", "json")

let perPagePrm = Some ("per_page", "10000")

let query = [formatPrm; perPagePrm] |> List.choose id

let response = Http.Request("http://api.worldbank.org/countries/chn;br/indicators/SP.POP.TOTL", query, silentHttpErrors = true)

let res =
    match response.Body with  
        | Text(json) -> 
            if response.StatusCode >= 400 then
                let doc = HtmlDocument.Parse(json)
                let body = doc.Body()
                let err = body.Descendants ["p"] 
                            |> Seq.map (fun v -> v.InnerText())
                            |>  String.concat "."
                printfn "%A" err
                [||]
            else
                let json = JsonValue.Parse(json).AsArray().[1]
                let countries = JsonConvert.DeserializeObject<IndicatorData[]>(json.ToString()) 
                countries
                //return XlTable.Create(countries, String.Empty, String.Empty, false, transposed, headers)
        | Binary(_) -> [||]


let frame = res |> Array.map (fun x -> x.Date, x.Country.Value, x.Value ) |> Frame.ofValues 

let dates = frame.RowKeys |> Seq.toArray

let x = frame.Rows.["2015"]
let y = x.TryGet("China")

let cols = frame.Columns

cols.Keys |> Seq.toList