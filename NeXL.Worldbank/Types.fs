namespace NeXL.Quandl
open NeXL.ManagedXll
open NeXL.XlInterop
open System
open Newtonsoft.Json
open Newtonsoft.Json.Linq

[<XlInvisible>]
type QuandlErrorMsg =
    {
     Code : string
     Message : string
    }


    
  