d:\humana\powershell\etl_sql2csv.ps1 -server "hulk", -db "adhocdata", -file "d:\humana\powershell\output\test.csv", -sql "select getdate() loadtime, * from adhocdata.sys.schemas ;"
