<?php

use Illuminate\Support\Facades\Route;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Log;
use Metadocx\Reporting\Converters\Excel\ExcelConverter;

Route::post("/Metadocx/Convert/Excel", function(Request $request) {

    $oConverter = new ExcelConverter();    
    $oConverter->loadOptions($request->input("ExportOptions"));    
    $sFileName = $oConverter->convert($request->input("ReportDefinition"));
    if ($sFileName !== false) {       
        $headers = ["Content-Type"=> "application/octet-stream"];
        return response()
                ->download($sFileName, "Report.xlsx", $headers)
                ->deleteFileAfterSend(true);
    } 
    

});