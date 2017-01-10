#r "../packages/DocumentFormat.OpenXml/lib/DocumentFormat.OpenXml.dll"
#r "../packages/NeXL/lib/net45/NeXL.XlInterop.dll"
#r "../packages/NeXL/lib/net45/NeXL.ManagedXll.dll"
#r "../packages/FSharp.Data/lib/net40/FSharp.Data.dll"
#r "../packages/Newtonsoft.Json/lib/net45/Newtonsoft.Json.dll"

open NeXL.ManagedXll
open NeXL.XlInterop
open System
open System.IO
open System.Runtime.InteropServices
open System.Data
open FSharp.Data
open FSharp.Data.JsonExtensions
open Newtonsoft.Json
open Newtonsoft.Json.Linq

[<XlInvisible>]
type IdValue =
    {
     Id : string
     Value : string
    }

[<XlInvisible>]
type CountryResponse =
    {
     Id : string
     Iso2Code : string
     Name : string
     Region : IdValue
     AdminRegion : IdValue
     IncomeLevel : IdValue
     LendingType : IdValue
     CapitalCity : string
     Longitude : Nullable<decimal>
     Latitude : Nullable<decimal>
    }

[<XlInvisible>]
type Country =
    {
     Name : string
     Id : string
     Iso2Code : string
     CapitalCity : string
     Longitude : Nullable<decimal>
     Latitude : Nullable<decimal>     
     Region : string
     AdminRegion : string
     IncomeLevel : string
     LendingType : string
    }

let toCountry (v : CountryResponse) : Country =
    {
        Name = v.Name
        Id = v.Id
        Iso2Code = v.Iso2Code
        CapitalCity = v.CapitalCity
        Longitude = v.Longitude
        Latitude = v.Latitude
        Region = v.Region.Value
        AdminRegion = v.AdminRegion.Value
        IncomeLevel = v.IncomeLevel.Value
        LendingType = v.LendingType.Value
    }

let formatPrm = Some ("format", "json")

let perPagePrm = Some ("per_page", "1000000")

let query = [formatPrm; perPagePrm] |> List.choose id

let response = Http.Request("http://api.worldbank.org/countries/dup", query, silentHttpErrors = true)

let res =
    match response.Body with  
        | Text(json) -> 
            if response.StatusCode >= 400 then
                //let err = JsonConvert.DeserializeObject<QuandlError>(json)
                printfn "%s" json
                [||]
            else
                printfn "%s" json
                let json = JsonValue.Parse(json).AsArray().[1]
                printfn "%s" (json.ToString())
                let countries = JsonConvert.DeserializeObject<CountryResponse[]>(json.ToString()) |> Array.map toCountry
                countries
                //return XlTable.Create(countries, String.Empty, String.Empty, false, transposed, headers)
        | Binary(_) -> [||]

response.Headers
