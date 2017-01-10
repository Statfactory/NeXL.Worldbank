namespace NeXL.Quandl
open NeXL.ManagedXll
open NeXL.XlInterop
open System
open System.IO
open System.Runtime.InteropServices
open System.Data
open FSharp.Data
open Newtonsoft.Json
open Newtonsoft.Json.Linq

[<XlQualifiedName(true)>]
module Worldbank =

    let private countriesUrl = "http://api.worldbank.org/countries"

    let private toCountry (v : CountryResponse) : Country =
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

    let getCountries(
                    [<XlArgHelp("Country code (optional, all countries will be returned if not specified)")>] countryCode : string option,
                    [<XlArgHelp("True if headers should be returned (optional, default is TRUE)")>] headers : bool option,
                    [<XlArgHelp("True if table should be returned as transposed (optional, default is FALSE)")>] transposed : bool option
                    ) =
        async  
            {
            let transposed = defaultArg transposed false

            let headers = defaultArg headers true

            let formatPrm = Some ("format", "json")

            let perPagePrm = Some ("per_page", "1000000")

            let query = [formatPrm; perPagePrm] |> List.choose id

            let countriesUrl =
                match countryCode with   
                    | Some(code) -> sprintf "%s/%s" countriesUrl code
                    | None -> countriesUrl

            let! response = Http.AsyncRequest(countriesUrl, query, silentHttpErrors = true)

            match response.Body with  
                | Text(json) -> 
                    if response.StatusCode >= 400 then
                        raise (new ArgumentException(json))
                        return XlTable.Empty
                    else
                        let json = JsonValue.Parse(json).AsArray()
                        if json.Length = 1 then
                            let err = JsonConvert.DeserializeObject<ErrorMessage>(json.[0].ToString())
                            raise (new ArgumentException(err.Message.[0].Key))
                            return XlTable.Empty
                        else
                            let countries = JsonConvert.DeserializeObject<CountryResponse[]>(json.[1].ToString()) |> Array.map toCountry
                            return XlTable.Create(countries, String.Empty, String.Empty, false, transposed, headers)
                | Binary(_) -> 
                    raise (new ArgumentException("Binary response received, json expected"))
                    return XlTable.Empty
             }

    let getErrors(newOnTop: bool) : IEvent<XlTable> =
        UdfErrorHandler.OnError |> Event.scan (fun s e -> e :: s) []
                                |> Event.map (fun errs ->
                                                  let errs = if newOnTop then errs |> List.toArray else errs |> List.rev |> List.toArray
                                                  XlTable.Create(errs, "", "", false, false, true)
                                             )


        
