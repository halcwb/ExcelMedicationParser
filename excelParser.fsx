

#I "./packages/ExcelProvider/lib"
#r "ExcelProvider.dll"

#I "./packages/FSharp.Data/lib/net40"
#r "FSharp.Data.dll"

#I "./packages/HtmlAgilityPack/lib/net40"
#r "HtmlAgilityPack.dll"

#r "Microsoft.Office.Interop.Excel"

#r "office"


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


[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module Assortment =

    type Assortment = ExcelFile<"template/AssortimentGPKTemplate_01.xlsx">

    type Drug = 
        {
            GPKcode : int
            ATCcode : string
            Name : string
            Label : string
            Deel : int
        }

    let createDrug gpk atc nm lb dl =
        {
            GPKcode = gpk
            ATCcode = atc
            Name = nm
            Label = lb
            Deel = dl
        }

    let parse (r: Assortment.Row) =
        try
            if r.GPKcode |> String.notNullOrEmpty then
                let gpk = 
                    let bl, code = r.GPKcode |> Int32.TryParse 
                    if bl then code else 0
                let atc = r.ATCCODE
                let nm = r.NMNM40
                let lb = r.Productnaam
                let dl = 
                    match r.DEELFACTOR with
                    | null -> 0
                    | _ ->
                        let bl, df = r.DEELFACTOR |> Int32.TryParse
                        if bl then df else 0
                if gpk = 0 |> not then createDrug gpk atc nm lb dl |> Some
                else None
            else None
        with
        | e ->
            printfn "Could not parse %A" r
            printfn "%s" e.Message
            None

    let get path =
        seq {
            for r in (new Assortment(path)).Data do
                match parse r with
                | Some r' -> yield r'
                | None -> ()
        }


[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module Prescription =

    open System
    open System.IO

    open FSharp.ExcelProvider

    type Prescription =
        {
            Department : string
            HospitalNumber : int
            LastName : string
            FirstName : string
            BirthDate : DateTime Option
            GestAgeWeeks : int Option
            GestAgeDays : int Option
            BirthWeight : float Option
            BirthWghtUnit : string
            Weight : float Option
            WeightUnit : string
            Start : DateTime option
            StartPrescriber : string
            Stop : DateTime Option
            StopPrescriber : string
            Generic : string
            Route : string
            Frequency : float Option
            FreqUnit : string
            Dose : float Option
            DoseUnit : string
            DoseTotal : float Option
            TotalUnit : string
            Text : string
        }
    
    type Prescriptions = ExcelFile<"template/PrescriptionTemplate_03.xlsx">

    let fromRow (p: Prescriptions.Row) =
        let getDate dt = 
            if dt = DateTime(1, 1, 1) then None else dt |> Some

        try 
            {
                Department = p.Department
                HospitalNumber = p.HospitalNumber |> int
                LastName = p.LastName
                FirstName = p.FirstName
                BirthDate = p.BirthDate |> getDate
                GestAgeWeeks = p.GestAgeWeeks |> int |> Some
                GestAgeDays = p.GestAgeDays |> int |> Some
                BirthWeight = p.BirthWeight |> Some
                BirthWghtUnit = p.BirthWghtUnit
                Weight = p.Weight |> Some
                WeightUnit = p.WeightUnit
                Start = p.StartDate |> getDate
                StartPrescriber = p.StartPrescriber
                Stop = p.StopDate |> getDate
                StopPrescriber = p.StopPrescriber
                Generic = p.Generic
                Route = p.Route
                Frequency = p.Frequency |> Some
                FreqUnit = p.FreqUnit
                Dose = p.Dose |> Some
                DoseUnit = p.DoseUnit
                DoseTotal = p.DoseTotal |> Some
                TotalUnit = p.TotalUnit
                Text = p.Text                    
            }
            |> Some
        with
        | e -> 
            printfn "Could not parse %A" p
            printfn "%s" e.Message
            None

    let get path = 
        seq { 
            for p in (new Prescriptions(path)).Data do
                match p |> fromRow with
                | Some p' -> yield p'
                | None -> ()
        }


[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module RouteMapping =
    
    open FSharp.ExcelProvider

    type RouteMapping = ExcelFile<"template/RouteMappingTemplate_01.xlsx">

    let get path = (new RouteMapping(path)).Data


[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module UnitMapping =
    
    open FSharp.ExcelProvider

    type UnitMapping = ExcelFile<"template/UnitMappingTemplate_01.xlsx">

    let get path = (new UnitMapping(path)).Data


module ExcelWriter =
    
    open System
    open System.IO

    open Microsoft.Office.Interop

    let toArray2D xs = 
        let rows = xs |> Seq.length
        let bln, cols = 
            xs 
            |> Seq.fold (fun (b, a) x ->
                let c = x |> Seq.length
                b || (a <> 0 && c <> a) ,
                if a = 0 || c < a then c else a  
            ) (false, 0)
        if bln then printfn "Warning: ragged array, data may be lost"
        Array2D.init rows cols (fun r c ->
            xs |> Seq.item r |> Seq.item c :> obj  
        ) :> obj


    let seqToExcel nm sh xs =
        let objs = new System.Collections.ArrayList()

        let xlApp = 
            let app = new Excel.ApplicationClass(Visible = true)
            objs.Add(app) |> ignore
            app

        let xlWorkBook = 
            let book = xlApp.Workbooks.Add()
            objs.Add(book) |> ignore
            book
    
        let xlSheet = 
            let sheet = xlWorkBook.Sheets.[1] :?> Excel.Worksheet
            objs.Add(sheet) |> ignore
            sheet

        xlSheet.Name <- sh

        let c1 = 
            let o = xlSheet.Cells.Item(1, 1)
            objs.Add(o) |> ignore
            o

        let c2 = 
            let o = xlSheet.Cells.Item(xs |> Seq.length, xs |> Seq.head |> Seq.length)
            objs.Add(o) |> ignore
            o

        let r = 
            let o = xlSheet.Range(c1, c2)
            objs.Add(o) |> ignore
            o

        r.Value2 <- xs |> toArray2D

        let path = Path.Combine(Environment.CurrentDirectory, nm + ".xlsx")

        if File.Exists(path) then File.Delete(path)
        xlWorkBook.SaveAs(path)

        xlWorkBook.Close(false)
        xlApp.Quit()
           
        GC.Collect()
        GC.WaitForPendingFinalizers()

        objs
        |> Seq.cast<obj>
        |> Seq.iter (Runtime.InteropServices.Marshal.FinalReleaseComObject >> ignore)


//let combine f p = IO.Path.Combine(p, f)
//
//let parent p =
//    let d = (new System.IO.DirectoryInfo(p))
//    d.Parent.FullName
//
//let path =
//    System.Environment.CurrentDirectory
//    |> parent
//    |> combine "GenPresCheck"
//    |> combine "NEO.xlsx"
//
//for p in Prescription.Prescriptions(path).Data do printfn "%A" (p.HospitalNumber)
//
//Prescription.get path
//|> Seq.filter(fun p -> 
//    p.Stop < p.Start
//)
