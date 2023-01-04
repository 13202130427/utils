<?php

namespace Uroad\Utils\File;


use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use Uroad\Utils\Common;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class Excel
{

    private static $header;
    private static $headerLine = [];

    private $fontName = 'Arial';
    private $fontSize = 20;
    private $fontBold = false;

    private static $type;
    private static $obj;
    private static $columnList = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
    private static $columnWidth = 20;
    private static $lineHigh = 30;
    private static $horizontal = Alignment::HORIZONTAL_CENTER;
    private static $vertical = Alignment::VERTICAL_CENTER;

    public static $instance = null;

    private static $setHeader = false;
    private static $setData = false;


    private function __construct($type,$config = [])
    {
        if (!empty($config)) {
            if (isset($config['font_name'])) $this->fontName = $config['font_name'];
            if (isset($config['font_size'])) $this->fontSize = $config['font_size'];
            if (isset($config['font_bold'])) $this->fontBold = $config['font_bold'];
            if (isset($config['column_width'])) self::$columnWidth = $config['column_width'];
            if (isset($config['line_high'])) self::$lineHigh = $config['line_high'];
            if (isset($config['horizontal'])) self::$horizontal = $config['horizontal'];
            if (isset($config['vertical'])) self::$vertical = $config['vertical'];
        }
        self::$type = $type;
        self::$obj = new Spreadsheet();
        if ($type == 1) {
            self::$obj->getDefaultStyle()->getFont()->setName($this->fontName);//设置字体
            self::$obj->getDefaultStyle()->getFont()->setSize($this->fontSize);//设置字体大小
            self::$obj->getDefaultStyle()->getFont()->setBold($this->fontBold);//设置是否加粗
            return;
        }
    }
    private function __clone()
    {
        // TODO: Implement __clone() method.
    }


    /**
     * 加载入口
     * @param array $config
     * @param int $type 0读 1写
     * @return Excel|null
     */
    public static function load($type = 0,$config = [])
    {
        if (!(self::$instance instanceof self)) {
            self::$instance = new self($type,$config);
        }
        return self::$instance;
    }

    /**
     * @param array $header 标题 支持单行多行 标准格式  [
     * [
     * 'A' => ['title'=>'','value' => '','column_width'=>'','line_high'=>'','merge_column'=>'B','merge_line'=>''],
     * 'C' => ['title'=>'','value' => '','column_width'=>'','merge_column'=>'','merge_line'=>''],
     * ],
     * [
     * 'A' => ['title'=>'','value' => '','column_width'=>'','merge_column'=>'','merge_line'=>''],
     * 'B' => ['title'=>'','value' => '','column_width'=>'','merge_column'=>'','merge_line'=>''],
     * 'C' => ['title'=>'','value' => '','column_width'=>'','merge_column'=>'','merge_line'=>'1'],
     * ]
     * ]
     * title：中文标题 value：英文标题 width 列宽 merge_column 合并列 merge_line 向上合并行数
     * @param int $sheetIndex 内置表 默认表一
     * @param string $sheetName 内置表名称 默认 Sheet0
     * @return bool|Excel
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setHeader($header = [],$sheetIndex = 0,$sheetName = '')
    {
        self::$headerLine[$sheetIndex] = 0;
        self::$obj->createSheet($sheetIndex);
        self::$obj->setActiveSheetIndex($sheetIndex);
        $sheet = self::$obj->getActiveSheet();
        if (empty($sheetName)) $sheetName = 'Sheet'.$sheetIndex;
        $sheet->setTitle($sheetName);
        $header = self::setHeaderFormat($header,$sheetIndex);
        if (!$header) return false;
        self::setHeaderStyle($header);
        if (self::$headerLine[$sheetIndex] == 0) return false;
        foreach ($header as $line => $head) {
            foreach ($head as $column => $data) {
                if (!Common::isEmpty($data,'value')) self::$header[$column] = $data['value'];
                if (!Common::isEmpty($data,'merge_line') && is_numeric($data['merge_line'])) {
                    //合并行
                    $lastLine = $line-$data['merge_line'];
                    if ($lastLine <= 0) $lastLine = 1;//超过条数默认第一行开始
                    $sheet->setCellValue($column.$lastLine, $data['title']);
                    continue;
                }
                $sheet->setCellValue($column.$line, $data['title']);
            }
        }
        return $this;
    }

    /**
     * 设置标题格式
     * @param $header
     * @param $sheetIndex
     * @return array|bool
     */
    private static function setHeaderFormat($header,$sheetIndex)
    {
        $data = [];
        try {
            //处理数据
            foreach ($header as $key => $value) {
                if (!is_array($value)) {
                    return false;
                }
                foreach ($value as $k =>$item) {
                    if (!is_array($item)) {
                        return false;
                    }
                    self::$headerLine[$sheetIndex] ++;
                    if (!in_array($k,self::$columnList)) {
                        $k = self::$columnList[$k];
                    }
                    //多行表头 标准格式
                    $data[self::$headerLine[$sheetIndex]][$k] = [
                        'title'=>$item['title'],
                        'value' => $item['value'],
                    ];
                    Common::isEmptySetValue($item,'column_width',$data[self::$headerLine[$sheetIndex]][$k],'column_width','');
                    Common::isEmptySetValue($item,'line_high',$data[self::$headerLine[$sheetIndex]][$k],'line_high','');
                    Common::isEmptySetValue($item,'merge_column',$data[self::$headerLine[$sheetIndex]][$k],'merge_column','');
                    Common::isEmptySetValue($item,'merge_line',$data[self::$headerLine[$sheetIndex]][$k],'merge_line','');
                }
            }
        }catch (\Throwable $exception) {
            return false;
        }
        return $data;
    }

    /**
     * 设置标题样式
     * @param $header
     */
    private static function setHeaderStyle($header)
    {
        $headerStyle = [];
        $sheet = self::$obj->getActiveSheet();
        foreach ($header as $line => $head) {
            foreach ($head as $column => $data) {
                Common::isEmptySetValue($data,'column_width',$headerStyle['column_width'],$column,self::$columnWidth);
                Common::isEmptySetValue($data,'line_high',$headerStyle['line_high'],$line,self::$lineHigh);
                Common::isEmptySetValue($data,'horizontal',$headerStyle['horizontal'],$column.$line,self::$horizontal);
                Common::isEmptySetValue($data,'vertical',$headerStyle['vertical'],$column.$line,self::$vertical);
                if(!Common::isEmpty($data,'merge_column') && in_array($data['merge_column'],self::$columnList)) {
                    //合并列
                    $sheet->mergeCells($column.$line.':'.$data['merge_column'].$line);
                }
                if (!Common::isEmpty($data,'merge_line') && is_numeric($data['merge_line'])) {
                    //合并行
                    $lastLine = $line-$data['merge_line'];
                    if ($lastLine <= 0) $lastLine = 1;//超过条数默认第一行开始
                    $sheet->mergeCells($column.$lastLine.':'.$column.$line);
                    Common::isEmptySetValue($data,'horizontal',$headerStyle['horizontal'],$column.$lastLine,self::$horizontal);
                    Common::isEmptySetValue($data,'vertical',$headerStyle['vertical'],$column.$lastLine,self::$vertical);
                    continue;
                }
            }
        }
        foreach ($headerStyle as $style => $data) {
            switch ($style) {
                case 'column_width':
                    foreach ($data as $key=>$value) {
                        $sheet->getColumnDimension($key)->setWidth($value);
                    }
                    break;
                case 'line_high':
                    foreach ($data as $key=>$value) {
                        $sheet->getRowDimension($key)->setRowHeight($value);
                    }
                    break;
                case 'horizontal':
                    foreach ($data as $key=>$value) {
                        $sheet->getStyle($key)->applyFromArray(['alignment'=>['horizontal'=>$value]]);
                    }
                    break;
                case 'vertical':
                    foreach ($data as $key=>$value) {
                        $sheet->getStyle($key)->applyFromArray(['alignment'=>['vertical'=>$value]]);
                    }
                    break;
            }
        }

    }

    /**
     * @param array $data 格式 [['value' => 'data']]
     * @param int $sheetIndex 内置表 默认表一
     * @return Excel
     */
    public function setData($data,$sheetIndex = 0)
    {
        self::$obj->setActiveSheetIndex($sheetIndex);
        $sheet = self::$obj->getActiveSheet();
        $startLine = self::$headerLine[$sheetIndex] +1;
        foreach ($data as $lineData) {
            foreach ($lineData as $column => $columnData) {
                $column = array_search($column,self::$header);
                $sheet->setCellValue( $column.$startLine, $columnData);
                $sheet->getStyle($column.$startLine)->applyFromArray(['alignment'=>[
                    'horizontal'=>self::$horizontal,
                    'vertical'=>self::$vertical
                ]]);
            }
            $startLine++;
        }
        return $this;
    }

    /**
     * 生成表格
     * @param string $path 路程 存在即保存指定路径
     * @param string $format 格式 Xlsx  Xls
     * @return bool|string
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function generate($path ='',$format = 'Xlsx')
    {
        if (self::$setData && self::$setHeader) return false;
        if ($format == 'Xlsx') {
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        } elseif ($format == 'Xls') {
            header('Content-Type: application/vnd.ms-excel');
        }
        self::$obj->setActiveSheetIndex(0);
        ob_end_clean();//清除缓冲区,避免乱码
        $objWriter = IOFactory::createWriter(self::$obj,$format);
        if (!empty($path)) {
            $objWriter->save($path);
            return $path;
        }
        $objWriter->save('php://output');
        exit;
    }

    public function dataConvert($header,$file,$format,$sheetIndex = 0) {
        try {
            $file = iconv("utf-8", "gb2312", $file);
            $objRead = IOFactory::createReader($format);
            if (empty($file) || !file_exists($file)) {
                throw new \Exception('文件不存在!');
            }
            if (!$objRead->canRead($file)) {
                throw new \Exception('格式错误');
            }
            //设置只读
            $objRead->setReadDataOnly(true);
            self::$obj = $objRead->load($file);
            $sheet = self::$obj->getSheet($sheetIndex);
            //获取总列数
            $columnAddress = $sheet->getHighestColumn();
            $columnNum = Coordinate::columnIndexFromString($columnAddress);
            //获取总行数
            $highestRow = $sheet->getHighestRow();
            /* 读取内容 */
            $data = [];
            //读取标题
            if (!self::matchHeader($header,$sheetIndex,$columnNum)) {
                throw new \Exception('表头不匹配');
            }
            for ($row = self::$headerLine[$sheetIndex]+1; $row <= $highestRow; $row++) {
                $isNull = true;
                for ($column = 1; $column <= $columnNum; $column++) {
                    $cellName = Coordinate::stringFromColumnIndex($column);
                    $data[$row][self::$header[$cellName]] = trim($sheet->getCell($cellName . $row)->getFormattedValue());
                    if (empty($data[$row][self::$header[$cellName]])) {
                        //是否被合并了
                        $result = $this->checkCellIsMerge($sheetIndex,$cellName . $row);
                        if ($result !== false) {
                            $data[$row][self::$header[$cellName]] = $result;
                        }
                    }
                    if (!empty($data[$row][self::$header[$cellName]])) {
                        $isNull = false;
                    }
                }
                /* 判断是否整行数据为空，是的话删除该行数据 */
                if ($isNull) {
                    unset($data[$row]);
                }
            }
            array_multisort($data);
            return $data;
        }catch (\Exception $exception) {
            throw new \Exception($exception);
        }
    }

    /**
     * 验证是否为合并单元格 是返回合并单元第一个赋值单元的值 不是返回false
     * @param $sheetIndex
     * @param $cell string 验证单元格 A1
     * @return bool|string
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function checkCellIsMerge($sheetIndex,$cell)
    {
        $cellColumn = array_search(substr($cell,0,1),self::$columnList);
        $cellRow = array_search(substr($cell,1),self::$columnList);
        $sheet = self::$obj->getSheet($sheetIndex);
        $mergeCells = $sheet->getMergeCells();
        foreach ($mergeCells as $mergeCell) {
            $mergeCellArr = explode(':',$mergeCell);
            $mergeStartColumn = array_search(substr($mergeCellArr[0],0,1),self::$columnList);
            $mergeEndColumn = array_search(substr($mergeCellArr[1],0,1),self::$columnList);
            $mergeStartRow = substr($mergeCellArr[0],1);
            $mergeEndRow = substr($mergeCellArr[1],1);
            //A1:B1 ||  A1:B2
            if ($cellColumn <= $mergeStartColumn && $cellColumn >= $mergeEndColumn && $cellRow >= $mergeStartRow && $cellRow <= $mergeEndRow) {
                return trim($sheet->getCell($mergeCellArr[0])->getFormattedValue());
            }
        }
        return false;
    }

    private function matchHeader($header,$sheetIndex,$columnNum)
    {
        //表头转换
        $header = self::headerConvert($header,$sheetIndex);
        if (!$header) return false;
        //获取文件对应表头数据
        $fileHeader = [];
        $sheet = self::$obj->getSheet($sheetIndex);
        for ($row = 1; $row <= self::$headerLine[$sheetIndex];$row++) {
            for ($column = 1; $column <= $columnNum; $column++) {
                $cellName = Coordinate::stringFromColumnIndex($column);
                $cellValue = trim($sheet->getCell($cellName . $row)->getFormattedValue());
                if (!empty($cellValue)) {
                    $fileHeader[$cellName] = $cellValue;
                }
            }
        }
        //文件表头与模板表头匹配
        foreach ($fileHeader as $column => $title) {
            if ($header[$column] != $title) {
                return false;
            }
        }
        return true;
    }

    /**
     * 表头转换
     * @param $header
     * @param $sheetIndex
     * @return bool
     */
    private function headerConvert($header,$sheetIndex)
    {
        self::$headerLine[$sheetIndex] = 0;
        $data = [];
        try {
            //处理数据
            foreach ($header as $key => $value) {
                if (!is_array($value)) {
                    return false;
                }
                self::$headerLine[$sheetIndex] ++;
                foreach ($value as $k =>$item) {
                    if (!is_array($item)) {
                        return false;
                    }
                    //多行表头 标准格式
                    $data[$k] = $item['title'];
                    self::$header[$k] = $item['value'];
                }
            }
        }catch (\Throwable $exception) {
            return false;
        }
        return $data;
    }


}