namespace NeXL.Worldbank
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
type IndicatorResponse =
    {
     Id : string
     Name : string
     Source : IdValue
     SourceNote : string
     SourceOrganization : string
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
     RegionId : string   
     Region : string
     AdminRegionId : string
     AdminRegion : string
     IncomeLevelId : string
     IncomeLevel : string
     LendingTypeId : string
     LendingType : string
    }

[<XlInvisible>]
type Indicator =
    {
     Id : string
     Name : string
     SourceId : string
     Source : string
     SourceNote : string
     SourceOrganization : string
    }

[<XlInvisible>]
type IndicatorData =
    {
     Indicator : IdValue
     Country : IdValue
     Value : Nullable<decimal>
     Decimal : Nullable<decimal>
     Date : string
    }

[<XlInvisible>]
type IncomeLevel =
    {
     IncomeLevelId : string
     IncomeLevel : string
    }

[<XlInvisible>]
type LendingType =
    {
     LendingTypeId : string
     LendingType : string
    }



    
  