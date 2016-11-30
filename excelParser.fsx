

#I "./packages/ExcelProvider/lib"
#r "ExcelProvider.dll"

#I "./packages/FSharp.Data/lib/net40"
#r "FSharp.Data.dll"

#I "./packages/HtmlAgilityPack/lib/net40"
#r "HtmlAgilityPack.dll"

#time

open System

open FSharp.ExcelProvider

Environment.CurrentDirectory <- __SOURCE_DIRECTORY__


[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module Double =

    let tryParse s =
        let (b, n) = Double.TryParse(s)
        if b then Some n else None

[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module Char =

    let letters = [|'a'..'z'|]

    let capitals = [|'A'..'Z'|]

    let isCapital c = capitals |> Seq.exists ((=) c)


[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module String =

    let apply f (s: string) = f s

    let nullOrEmpty = apply String.IsNullOrEmpty

    let notNullOrEmpty = nullOrEmpty >> not

    let splitAt s1 (s2: string) =
        let cs = s1 |> Array.ofSeq
        s2.Split(cs)

    let arrayConcat (cs : char[]) = String.Concat(cs)

// GPKcode	Productnaam	NMNM40	STATUS	ATCCODE



[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module Assortment =

    type Assortment = ExcelFile<"AssortimentGPK.xlsx">

    type Drug = 
        {
            GPKcode : int
            ATCcode : string
            Name : string
            Label : string
            Status : string
        }

    let createDrug gpk atc nm lb st =
        {
            GPKcode = gpk
            ATCcode = atc
            Name = nm
            Label = lb
            Status = st
        }


    let parse () =
        [
            for r in (new Assortment()).Data do
                if r.GPKcode |> String.notNullOrEmpty then
                    let gpk = r.GPKcode |> Int32.Parse
                    let atc = r.ATCCODE
                    let nm = r.NMNM40
                    let lb = r.Productnaam
                    let st = r.STATUS
                    yield createDrug gpk atc nm lb st
        ]




