namespace NeXL.Worldbank
open NeXL.ManagedXll
open NeXL.XlInterop
open System
open System.IO
open System.Runtime.InteropServices
open System.Data
open FSharp.Data
open Newtonsoft.Json
open Newtonsoft.Json.Linq
open Deedle

[<XlQualifiedName(true)>]
module Worldbank =

    let private frameToDataTable (frame : Frame<'T, string>) : DataTable =
        let dbTable = new DataTable()
        let dateCol = new DataColumn("Date", typeof<'T>)
        let colNames = frame.Columns.Keys |> Seq.toArray
        let cols = colNames |> Array.map (fun colName -> new DataColumn(colName, typeof<decimal>)) 
        dbTable.Columns.Add(dateCol)
        dbTable.Columns.AddRange(cols)
        frame.RowKeys |> Seq.iter (fun date ->
                                       let frameRow = frame.Rows.[date]
                                       let dbRow = dbTable.NewRow()
                                       dbRow.["Date"] <- date
                                       colNames |> Array.iter (fun colName ->
                                                                   let v = frameRow.TryGet(colName)
                                                                   if v.HasValue then
                                                                       dbRow.[colName] <- v.Value
                                                                   else
                                                                       dbRow.[colName] <- DBNull.Value
                                                              )
                                       dbTable.Rows.Add(dbRow)
                                  )
        dbTable

    let private countriesUrl = "http://api.worldbank.org/countries"
    let private incomeLevelsUrl = "http://api.worldbank.org/incomeLevels"
    let private lendingTypesUrl = "http://api.worldbank.org/lendingTypes"
    let private indicatorsUrl = "http://api.worldbank.org/indicators"

    let private getIndicatorDataUrl countries indicator = sprintf "http://api.worldbank.org/countries/%s/indicators/%s" countries indicator

    let private toCountry (v : CountryResponse) : Country =
        {
         Name = v.Name
         Id = v.Id
         Iso2Code = v.Iso2Code
         CapitalCity = v.CapitalCity
         Longitude = v.Longitude
         Latitude = v.Latitude
         RegionId = v.Region.Id
         Region = v.Region.Value
         AdminRegionId = v.AdminRegion.Id
         AdminRegion = v.AdminRegion.Value
         IncomeLevelId = v.IncomeLevel.Id
         IncomeLevel = v.IncomeLevel.Value
         LendingTypeId = v.LendingType.Id
         LendingType = v.LendingType.Value
        }

    let private toIndicator (v : IndicatorResponse) : Indicator =
        {
        Id = v.Id
        Name = v.Name
        SourceId = v.Source.Id
        Source = v.Source.Value
        SourceNote = v.SourceNote
        SourceOrganization = v.SourceOrganization
        }

    let private toIncomeLevel (v : IdValue) : IncomeLevel =
        {
            IncomeLevelId = v.Id
            IncomeLevel = v.Value
        }

    let private toLendingType (v : IdValue) : LendingType =
        {
            LendingTypeId = v.Id
            LendingType = v.Value
        }

    let getCountries(
                    [<XlArgHelp("Country Ids (optional, semicolon delimited list or a row/column range)")>] countryIds : string[] option,
                    [<XlArgHelp("Region Ids (optional, semicolon delimited list or a row/column range)")>] regionIds : string[] option,
                    [<XlArgHelp("IncomeLevel Ids (optional, semicolon delimited list or a row/column range)")>] incomeLevelIds : string[] option,
                    [<XlArgHelp("LendingType Ids (optional, semicolon delimited list or a row/column range)")>] lendingTypeIds : string[] option,
                    [<XlArgHelp("True if headers should be returned (optional, default is TRUE)")>] headers : bool option,
                    [<XlArgHelp("True if table should be returned as transposed (optional, default is FALSE)")>] transposed : bool option
                    ) =
        async  
            {
            let transposed = defaultArg transposed false

            let headers = defaultArg headers true

            let formatPrm = Some ("format", "json")

            let perPagePrm = Some ("per_page", "1000000")

            let regionIdsPrm = regionIds |> Option.map (fun ids -> "region", ids |> String.concat ";")

            let incomeLevelIdsPrm = incomeLevelIds |> Option.map (fun ids -> "incomeLevel", ids |> String.concat ";")

            let lendingTypeIdsPrm = lendingTypeIds |> Option.map (fun ids -> "lendingType", ids |> String.concat ";")

            let query = [formatPrm; perPagePrm; regionIdsPrm; incomeLevelIdsPrm; lendingTypeIdsPrm] |> List.choose id

            let countriesUrl =
                match countryIds with   
                    | Some(ids) -> sprintf "%s/%s" countriesUrl (ids |> String.concat ";")
                    | None -> countriesUrl

            let! response = Http.AsyncRequest(countriesUrl, query, silentHttpErrors = true)

            match response.Body with  
                | Text(json) -> 
                    if response.StatusCode >= 400 then
                        let doc = HtmlDocument.Parse(json)
                        let body = doc.Body()
                        let err = body.Descendants ["p"] 
                                    |> Seq.map (fun v -> v.InnerText())
                                    |>  String.concat "."
                        raise (new ArgumentException(err))
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

    let getIncomeLevels(
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

            let! response = Http.AsyncRequest(incomeLevelsUrl, query, silentHttpErrors = true)

            match response.Body with  
                | Text(json) -> 
                    if response.StatusCode >= 400 then
                        let doc = HtmlDocument.Parse(json)
                        let body = doc.Body()
                        let err = body.Descendants ["p"] 
                                    |> Seq.map (fun v -> v.InnerText())
                                    |>  String.concat "."
                        raise (new ArgumentException(err))
                        return XlTable.Empty
                    else
                        let json = JsonValue.Parse(json).AsArray()
                        if json.Length = 1 then
                            let err = JsonConvert.DeserializeObject<ErrorMessage>(json.[0].ToString())
                            raise (new ArgumentException(err.Message.[0].Key))
                            return XlTable.Empty
                        else
                            let incLevels = JsonConvert.DeserializeObject<IdValue[]>(json.[1].ToString()) |> Array.map toIncomeLevel
                            return XlTable.Create(incLevels, String.Empty, String.Empty, false, transposed, headers)
                | Binary(_) -> 
                    raise (new ArgumentException("Binary response received, json expected"))
                    return XlTable.Empty
             }

    let getLendingTypes(
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

            let! response = Http.AsyncRequest(lendingTypesUrl, query, silentHttpErrors = true)

            match response.Body with  
                | Text(json) -> 
                    if response.StatusCode >= 400 then
                        let doc = HtmlDocument.Parse(json)
                        let body = doc.Body()
                        let err = body.Descendants ["p"] 
                                    |> Seq.map (fun v -> v.InnerText())
                                    |>  String.concat "."
                        raise (new ArgumentException(err))
                        return XlTable.Empty
                    else
                        let json = JsonValue.Parse(json).AsArray()
                        if json.Length = 1 then
                            let err = JsonConvert.DeserializeObject<ErrorMessage>(json.[0].ToString())
                            raise (new ArgumentException(err.Message.[0].Key))
                            return XlTable.Empty
                        else
                            let lendTypes = JsonConvert.DeserializeObject<IdValue[]>(json.[1].ToString()) |> Array.map toLendingType
                            return XlTable.Create(lendTypes, String.Empty, String.Empty, false, transposed, headers)
                | Binary(_) -> 
                    raise (new ArgumentException("Binary response received, json expected"))
                    return XlTable.Empty
             }

    let getIndicators(
                        [<XlArgHelp("Indicator Ids (optional, semicolon delimited list or a row/column range)")>] indicatorIds : string[] option,
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

            let indicatorsUrl =
                match indicatorIds with   
                    | Some(ids) -> sprintf "%s/%s" indicatorsUrl (ids |> String.concat ";")
                    | None -> indicatorsUrl

            let! response = Http.AsyncRequest(indicatorsUrl, query, silentHttpErrors = true)

            match response.Body with  
                | Text(json) -> 
                    if response.StatusCode >= 400 then
                        let doc = HtmlDocument.Parse(json)
                        let body = doc.Body()
                        let err = body.Descendants ["p"] 
                                    |> Seq.map (fun v -> v.InnerText())
                                    |>  String.concat "."
                        raise (new ArgumentException(err))
                        return XlTable.Empty
                    else
                        let json = JsonValue.Parse(json).AsArray()
                        if json.Length = 1 then
                            let err = JsonConvert.DeserializeObject<ErrorMessage>(json.[0].ToString())
                            raise (new ArgumentException(err.Message.[0].Key))
                            return XlTable.Empty
                        else
                            let indicators = JsonConvert.DeserializeObject<IndicatorResponse[]>(json.[1].ToString()) |> Array.map toIndicator
                            return XlTable.Create(indicators, String.Empty, String.Empty, false, transposed, headers)
                | Binary(_) -> 
                    raise (new ArgumentException("Binary response received, json expected"))
                    return XlTable.Empty
             }

    let getIndicatorData(
                         [<XlArgHelp("Country Ids (semicolon delimited list or a row/column range)")>] countryIds : string[],
                         [<XlArgHelp("Indicator Id")>] indicatorId : string,
                         [<XlArgHelp("True if headers should be returned (optional, default is TRUE)")>] headers : bool option,
                         [<XlArgHelp("True if table should be returned as transposed (optional, default is FALSE)")>] transposed : bool option
                         ) =
        async  
            {
            let transposed = defaultArg transposed false

            let headers = defaultArg headers true

            let formatPrm = Some ("format", "json")

            let perPagePrm = Some ("per_page", "20000")

            let query = [formatPrm; perPagePrm] |> List.choose id

            let indicatorDataUrl = getIndicatorDataUrl (countryIds |> String.concat ";") indicatorId

            let! response = Http.AsyncRequest(indicatorDataUrl, query, silentHttpErrors = true)

            match response.Body with  
                | Text(json) -> 
                    if response.StatusCode >= 400 then
                        let doc = HtmlDocument.Parse(json)
                        let body = doc.Body()
                        let err = body.Descendants ["p"] 
                                    |> Seq.map (fun v -> v.InnerText())
                                    |>  String.concat "."
                        raise (new ArgumentException(err))
                        return XlTable.Empty
                    else
                        let json = JsonValue.Parse(json).AsArray()
                        if json.Length = 1 then
                            let err = JsonConvert.DeserializeObject<ErrorMessage>(json.[0].ToString())
                            raise (new ArgumentException(err.Message.[0].Key))
                            return XlTable.Empty
                        else
                            let indicatorData = JsonConvert.DeserializeObject<IndicatorData[]>(json.[1].ToString()) 
                                                    |> Array.map (fun x -> x.Date, x.Country.Value, x.Value )
                                                    |> Frame.ofValues 
                                                    |> frameToDataTable
                            return XlTable(indicatorData, String.Empty, String.Empty, false, transposed, headers)
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


        
