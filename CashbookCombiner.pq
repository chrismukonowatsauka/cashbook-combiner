// Cashbook Combiner - Power Query (M)
// Stacks every bank sheet of a multi-sheet cashbook into one table.
// Author: Chris Mukonowatsauka - linkedin.com/in/cmkn
// Setup: name a cell CashbookPath holding the cashbook's full path, then paste this into a blank query.

let
    // ---------- settings ----------
    CashbookPath = Text.Trim(Text.From(Excel.CurrentWorkbook(){[Name = "CashbookPath"]}[Content]{0}[Column1])),
    SkipSheets   = {"Org", "Bank charges & Imtt"},     // non-bank sheets to ignore

    // ---------- helpers ----------
    Norm = (v as any) as text =>
        if v = null then "" else Text.Upper(Text.Select(Text.From(v), {"A".."Z", "a".."z", "0".."9"})),
    Txt = (v as any) =>
        if v = null then null else let t = Text.Trim(Text.From(v)) in if t = "" then null else t,
    Num = (v as any) as number =>
        let x = try Number.From(if v is text then Text.Remove(v, ",") else v) otherwise null
        in if x = null then 0 else x,
    NumOrNull = (v as any) =>
        try Number.From(if v is text then Text.Remove(v, ",") else v) otherwise null,
    // real Excel dates, or text dates typed as day/month/year (e.g. "2/10/2026 02:00:00")
    ToDate = (v as any) =>
        if v is datetime then Date.From(v)
        else if v is date then v
        else if v is text then
            try
                let p = Text.Split(Text.Trim(v), "/")
                in #date(Number.From(Text.Start(p{2}, 4)), Number.From(p{1}), Number.From(p{0}))
            otherwise null
        else null,

    // ---------- cleans one bank sheet (handles both layouts) ----------
    //  Balance = the running balance already worked out on the bank sheet
    //  Layout A: Date | PostRef | StmtRef | Description | Branch | DR | CR | Balance
    //  Layout B: Date | PostRef | Description | Branch | DR | CR | Balance
    CleanSheet = (t as table) as table =>
        let
            t0   = Table.ReplaceErrorValues(t, List.Transform(Table.ColumnNames(t), each {_, null})),
            Recs = List.Buffer(Table.ToRecords(t0)),
            Col  = (r as record, i as number) => Record.FieldOrDefault(r, "Column" & Text.From(i), null),
            PosA = List.PositionOf(List.Transform(Recs, (r) => Text.StartsWith(Norm(Col(r, 4)), "OPENINGBAL")), true),
            PosB = List.PositionOf(List.Transform(Recs, (r) => Text.StartsWith(Norm(Col(r, 3)), "OPENINGBAL")), true),
            Pos  = if PosA < 0 then PosB else if PosB < 0 then PosA else List.Min({PosA, PosB}),
            Off  = if Pos = PosA then 0 else -1,
            Body = if Pos < 0 then {} else List.Skip(Recs, Pos + 1),
            Mapped = List.Transform(Body, (r) => [
                Date             = ToDate(Col(r, 1)),
                #"Posting Ref"   = Txt(Col(r, 2)),
                #"Statement Ref" = if Off = 0 then Txt(Col(r, 3)) else null,
                Description      = Txt(Col(r, 4 + Off)),
                Branch           = Txt(Col(r, 5 + Off)),
                DR               = Num(Col(r, 6 + Off)),
                CR               = Num(Col(r, 7 + Off)),
                Balance          = NumOrNull(Col(r, 8 + Off))
            ]),
            // the opening balance line, so each bank's running balance starts from it
            Opening = if Pos < 0 then {} else {[
                Date             = ToDate(Col(Recs{Pos}, 1)),
                #"Posting Ref"   = null,
                #"Statement Ref" = null,
                Description      = "Opening Balance",
                Branch           = null,
                DR               = 0,
                CR               = 0,
                Balance          = NumOrNull(Col(Recs{Pos}, 8 + Off))
            ]},
            // keep every line that has a date and an amount - with or without a branch
            Kept = List.Select(Mapped, (m) => m[Date] <> null and (m[DR] <> 0 or m[CR] <> 0))
        in
            Table.FromRecords(Opening & Kept, type table [Date = nullable date, #"Posting Ref" = nullable text,
                #"Statement Ref" = nullable text, Description = nullable text,
                Branch = nullable text, DR = number, CR = number, Balance = nullable number]),

    // ---------- read the cashbook ----------
    Source   = Excel.Workbook(File.Contents(CashbookPath), null, true),
    Sheets   = Table.SelectRows(Source, each [Kind] = "Sheet" and not List.Contains(SkipSheets, Text.Trim([Name]))),
    WithRows = Table.AddColumn(Sheets, "Rows", each CleanSheet([Data])),
    Slim     = Table.SelectColumns(WithRows, {"Name", "Rows"}),
    Expanded = Table.ExpandTableColumn(Slim, "Rows",
                 {"Date", "Posting Ref", "Statement Ref", "Description", "Branch", "DR", "CR", "Balance"}),
    BankName = Table.TransformColumns(Table.RenameColumns(Expanded, {{"Name", "Bank Sheet"}}), {{"Bank Sheet", Text.Trim}}),

    // ---------- extra columns ----------
    AddCur = Table.AddColumn(BankName, "Currency", each
                let n = Text.Upper([Bank Sheet]) in
                if Text.Contains(n, "USD") then "USD"
                else if Text.Contains(n, "ZAR") then "ZAR"
                else if Text.Contains(n, "ZWL") or Text.Contains(n, "ZWG") then "ZWG"
                else null, type nullable text),
    AddStatus = Table.AddColumn(AddCur, "Status", each
                if [Description] = "Opening Balance" and [Posting Ref] = null then "Opening"
                else if [Posting Ref] = null then "No ref"
                else if Text.Upper([Posting Ref]) = "X" then "X"
                else "Posted", type text),
    AddNet  = Table.AddColumn(AddStatus, "Net", each [CR] - [DR], type number),
    Typed   = Table.TransformColumnTypes(AddNet, {{"Date", type date}, {"DR", type number}, {"CR", type number}, {"Balance", type number}}),
    Ordered = Table.SelectColumns(Typed, {"Date", "Bank Sheet", "Currency", "Posting Ref", "Status",
                "Statement Ref", "Description", "Branch", "DR", "CR", "Net", "Balance"})
    // rows stay in the same order as each bank sheet (no re-sorting)
in
    Ordered
