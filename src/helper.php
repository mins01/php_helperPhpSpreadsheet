<?
/**
 * 엑셀 다운로드
 * phpspreadsheet 사용
 */
function downloadXlsx($fn=null,$conf=null,$body=null,$header=null,$footer=null){

    if(!isset($fn[0])) $fn = '엑셀다운로드';

    //-- STYLE
    $styleBoldRed = [
        'font' => [
            'bold' => true,
            'color'=>[
                'argb'=>PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED,
            ]
        ],
    ];
    $styleCenterCenter = [
        'alignment' => [
            'horizontal' => PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            'vertical' => PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        ],
    ];
    $styleAllBorders = [
        'borders' => [
            'allBorders' => [
                // 'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                'color' => ['argb' => '00000000'],
            ],
        ],
    ];
    $styleLink = [
        'font'  => array(
            'color' => array('rgb' => '0000FF'),
            'underline' => 'single',
        )
    ];
    $styleHeadFill = [
        'fill'=>[
            'fillType'=>\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
            'color' => array('rgb' => 'C6EFCE'),
        ]
    ];

    $styleHeader = $styleCenterCenter+$styleAllBorders+$styleHeadFill;
    $styleRows = $styleAllBorders;
    $styleFooter = $styleAllBorders+$styleBoldRed;
    

    $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
    $spreadsheet->getProperties()->setCreator('RAPS')
        ->setLastModifiedBy('RAPS')
        ->setTitle('RAPS')
        ->setSubject('RAPS')
        ->setDescription('RAPS')
        ->setKeywords('RAPS')
        ->setCategory('RAPS')
        ->setDescription(
            "CREATE IN RAPS"
        );

    //--- 첫 시트
    $sheet = $spreadsheet->getActiveSheet();
    $default_width = 20;//20글자
    $sheet->getDefaultColumnDimension()->setWidth($default_width); // 각 시트에 기본 너비 설정함
    // $sheet->getDefaultColumnDimension()->setAutoSize(true);
    // $sheet->setCellValue('A1', '테스트입니다.');

    //--- column 설정
    $cf = $conf['sheet']??null;
    if(isset($cf['columns'])){
        foreach($cf['columns'] as $k => $column){
            $cIdx = $k+1; //column idx
            if(isset($column['width'])){
                if($column['width'] =='auto'){
                    $sheet->getColumnDimensionByColumn($cIdx)->setAutoSize(true);
                }else{
                    $sheet->getColumnDimensionByColumn($cIdx)->setWidth($column['width']);
                }
            }            
        }
    }



    $rIdx = 1; // row idx
    if(isset($header[0][0])){
        $d = & $header;
        $firstCellCoord = 'A'.$rIdx;
        $sheet->fromArray($d,null,'A'.$rIdx);
        $rIdx+= count($d);
        $lastColAlpha = $sheet->getHighestColumn($rIdx-1);
        $lastCellCoord = $lastColAlpha.($rIdx-1);
        $sheet->getStyle("{$firstCellCoord}:{$lastCellCoord}")->applyFromArray($styleHeader);
        //-- 스타일
        $cf = $conf['header']??null;
        if($cf && isset($cf['columns'])){
            foreach($cf['columns'] as $k => $column ){
                if(!isset($column)) continue;
                $alpha = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($k+1);
                $colCoord = $alpha.($rIdx - count($d)).':'.$alpha.($rIdx - 1);
                if(isset($column['style'])){
                    $sheet->getStyle($colCoord)->applyFromArray($column['style']);
                }
            }
        }

        unset($d);
    }
    if(isset($body[0][0])){
        //-- 값
        $d = & $body;
        $firstCellCoord = 'A'.$rIdx;
        $sheet->fromArray($d,null,'A'.$rIdx);
        $rIdx+= count($d);
        $lastColAlpha = $sheet->getHighestColumn($rIdx-1);
        $lastCellCoord = $lastColAlpha.($rIdx-1);
        $sheet->getStyle("{$firstCellCoord}:{$lastCellCoord}")->applyFromArray($styleRows);
        //-- 스타일
        $cf = $conf['body']??null;
        if($cf && isset($cf['columns'])){
            foreach($cf['columns'] as $k => $column ){
                if(!isset($column)) continue;
                $alpha = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($k+1);
                $colCoord = $alpha.($rIdx - count($d)).':'.$alpha.($rIdx - 1);
                if(isset($column['style'])){
                    $sheet->getStyle($colCoord)->applyFromArray($column['style']);
                }
            }
        }

        unset($d);
        
    }
    if(isset($footer[0][0])){
        $d = & $footer;
        $firstCellCoord = 'A'.$rIdx;
        $sheet->fromArray($d,null,'A'.$rIdx);
        $rIdx+= count($d);
        $lastColAlpha = $sheet->getHighestColumn($rIdx-1);
        $lastCellCoord = $lastColAlpha.($rIdx-1);
        $sheet->getStyle("{$firstCellCoord}:{$lastCellCoord}")->applyFromArray($styleFooter);
        //-- 스타일
        $cf = $conf['footer']??null;
        if($cf && isset($cf['columns'])){
            foreach($cf['columns'] as $k => $column ){
                if(!isset($column)) continue;
                $alpha = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($k+1);
                $colCoord = $alpha.($rIdx - count($d)).':'.$alpha.($rIdx - 1);
                if(isset($column['style'])){
                    $sheet->getStyle($colCoord)->applyFromArray($column['style']);
                }
            }
        }
        unset($d);
    }
    $sheet->setSelectedCell('A1');
    

    // //--- col 설정
    // $ri = 1; //row idx
    // $ci = 1; //cell idx
    
    // // print_r($conf['fields']);exit;
    // if(isset($conf['fields'])){
    //     foreach($conf['fields'] as $field){
    //         // $ca = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($ci);
    //         if(isset($field['width']) && $field['width'] !=='auto'){
    //             if($field['width'] =='auto'){
    //                 $sheet->getColumnDimensionByColumn($ci)->setAutoSize(true);
    //             }else{
    //                 $sheet->getColumnDimensionByColumn($ci)->setWidth($field['width']);
    //             }
    //         }            
    //         $cell = $sheet->getCell([$ci,$ri]);
    //         $cell->setValue($field['value']);
    //         $cell->getStyle()->applyFromArray($styleCenterCenter+$styleAllBorders+$styleHeadFill);
    //         $ci++;
    //     }
    //     $ri++;
    // }
    // foreach($rows as $row){
    //     $ci = 1; //cell idx
    //     foreach($row as $k => $v){
    //         $field = $conf['fields'][$k]??null;
    //         $cell = $sheet->getCell([$ci,$ri]);
    //         $cell->setValue($v);
    //         $cell->getStyle()->applyFromArray($styleCenterCenter+$styleAllBorders);
    //         $cell->getStyle()->getNumberFormat()
    //         ->setFormatCode(
    //             PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT
    //         );
    //         $ci++;
    //     }
    //     $ri++;
    // }




    // 출력부 - xlsx
    $writer = new PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="'.$fn.'.xlsx"');
    header('Cache-Control: max-age=0');
    $writer->save('php://output');
exit;
    // 테스트 출력부 - HTML
    $writer = new PhpOffice\PhpSpreadsheet\Writer\Html($spreadsheet);    
    $writer->save('php://output');


    exit();
}