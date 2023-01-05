<?php 
namespace Metadocx\Reporting\Converters\Excel;

use Illuminate\Support\Facades\Storage;
use Illuminate\Support\Facades\Log;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class ExcelConverter {

    protected $_sOutputFileName = null;
    protected $_aOptions = [];
    
    public function convert($aReportDefinition) {
        
        $this->_sOutputFileName = storage_path("app/" . uniqid("Word") . ".xlsx");

        $spreadsheet = new Spreadsheet();

        $sheet = $spreadsheet->getActiveSheet();
        
        foreach($aReportDefinition["sections"] as $aSection) {

            $colIndex = 1;
            $sheet->setTitle($aSection["properties"]["name"]);

            if (array_key_exists("model", $aSection)) {
                foreach($aSection["model"] as $aColumn) {
                    $bVisible = true;
                    if (array_key_exists("visible", $aColumn) && $this->toBool($aColumn["visible"]) === false) {
                        $bVisible = false;
                    }
                    if ($bVisible) {
                        $colKey = Coordinate::stringFromColumnIndex($colIndex);
                        $sheet->getColumnDimension($colKey)->setAutoSize(true);
                        $cell = $sheet->getCellByColumnAndRow($colIndex, 1);
                        $cell->getStyle()->getFill()->setFillType(Fill::FILL_SOLID);
                        $cell->getStyle()->getFill()->getStartColor()->setARGB("FFC0C0C0");
                        $cell->getStyle()->getFont()->setBold(true);
                        $cell->setValue($aColumn["label"]);
                        
                        $sheet->getStyle($colKey .":" . $colKey)->getAlignment()->setWrapText(true);

                        switch ($aColumn["type"]) {
                            case "string":
                                $sheet->getStyle($colKey .":" . $colKey)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_TEXT);
                                break;
                            case "number":
                                $sheet->getStyle($colKey .":" . $colKey)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00);
                                $sheet->getStyle($colKey .":" . $colKey)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
                                break;
                        }
                        

                        $colIndex++;
                    }
                }
                
                $nRowIndex = 2;

                foreach($aSection["data"] as $aRow) {
                                                        
                    if (array_key_exists("__visible", $aRow) && $this->toBool($aRow["__visible"]) === false) {
                        Log::debug("SKIPPING ROW");
                        continue;
                    }

                    $colIndex = 1;
                    foreach($aSection["model"] as $aColumn) {
                        $bVisible = true;
                        if (array_key_exists("visible", $aColumn) && $this->toBool($aColumn["visible"]) === false) {
                            $bVisible = false;
                        }
                        if ($bVisible) {
                            $cell = $sheet->getCellByColumnAndRow($colIndex, $nRowIndex);
                            $cell->setValue($aRow[$aColumn["name"]]);
                            $colIndex++;
                        }
                    }

                    $nRowIndex++;

                }

            }

        }



        $writer = new Xlsx($spreadsheet);
        $writer->save($this->_sOutputFileName);

        return $this->_sOutputFileName;

    }

    public function loadOptions($options) {

        $this->_aOptions = [];

    }

    public function __get(string $name) {
        $name = str_replace("_", "-", $name);
        return $this->_aOptions[$name];
    }

    public function __set(string $name, mixed $value) {
        $name = str_replace("_", "-", $name);
        $this->_aOptions[$name] = $value;
        return $this;
    }

    private function toBool($value) {
        if (is_bool($value)) {
            return (bool) $value;
        }

        $value = strtolower(trim($value));

        $aTrueValues = ["1","y","o","yes","true","oui","vrai","on","checked",true,1];
        
        if (in_array($value, $aTrueValues, true)) {
            return true;
        } else {
            return false;
        }
    }

}