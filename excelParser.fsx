

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


    let parse path =
        [
            for r in (new Assortment(path)).Data do
                if r.GPKcode |> String.notNullOrEmpty then
                    let gpk = r.GPKcode |> Int32.Parse
                    let atc = r.ATCCODE
                    let nm = r.NMNM40
                    let lb = r.Productnaam
                    let st = r.STATUS
                    yield createDrug gpk atc nm lb st
        ]


module Prescription =

    open System
    open System.IO

    open FSharp.ExcelProvider
    
    type Prescription = ExcelFile<"prescriptions.xlsx">

    let get path = (new Prescription(path)).Data




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

//        xs
//        |> Seq.iteri (fun i xs' ->
//            xs' 
//            |> Seq.iteri (fun j x -> 
//                let cell = xlSheet.Cells.Item(i + 1, j + 1) :?> Excel.Range
//                cell.Item(1, 1) <- (x :> obj)
//                objs.Add(cell) |> ignore)
//        )

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


