namespace NeXL.Quandl
open NeXL.ManagedXll
open NeXL.XlInterop
open System
open Newtonsoft.Json
open Newtonsoft.Json.Linq

[<XlInvisible>]
type IdKeyValue =
    {
     Id : string
     Key : string
     Value : string
    }

[<XlInvisible>]
type ErrorMessage =
    {
        Message : IdKeyValue[]
    }

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



    
  