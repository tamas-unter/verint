function Alarms-Table{
    $alarms_file= 'C:\Users\tpasztor\OneDrive - Verint Systems Ltd\ps\Copy of WFO_V15_2_Alarms_Spreadsheet.xlsx'
    $xl=New-Object -ComObject Excel.Application
    $wb=$xl.Workbooks.Open($alarms_file)
    $sheet=$wb.Sheets[2]
    # read the column headers
    $cols=@();$i=0;$v=$true;
    while(($i++ -lt $sheet.Columns.Count) -and $v) {
        $col=$sheet.Columns.Item($i).Rows.Item(1).Text;
        $v=($col.Length -gt 0);
        if($v){$cols+=$col}
    }
    # read contents
    $i=0;$v=$true;
    while(($i++ -lt $sheet.Rows.Count) -and $v){
        $v=($sheet.Columns.Item(1).Rows.Item($i).Text.Length -gt 0)
    };
    $lastrow=$i-2
    $alarms=for($row=1;$row -lt $lastrow;$row++){
        [pscustomobject]$a=@{};
        for($i=0;$i -lt $cols.Count;$i++){
            $a+=@{$cols[$i]=$sheet.Rows.Item($row).Columns.Item($i+1).Text}
        };
        $a
    }
    $wb.Close()
    $xl.Quit()
    $alarms
}