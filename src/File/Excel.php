<?php

namespace Uroad\Utils\File;



use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Excel
{

    private static $header;
    private static $headerLine = 0;

    private $fontName = 'Arial';
    private $fontSize = 20;
    private $fontBold = false;


    private static $obj;
    private static $columnList = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
    private static $columnWidth = 200;

    public static $instance = null;

    private static $setHeader = false;
    private static $setData = false;

    private function __construct($config = [])
    {
        if (!empty($config)) {
            if (isset($config['font_name'])) $this->fontName = $config['font_name'];
            if (isset($config['font_size'])) $this->fontSize = $config['font_size'];
            if (isset($config['font_bold'])) $this->fontBold = $config['font_bold'];
        }
        self::$obj = new Spreadsheet();
        self::$obj->getDefaultStyle()->getFont()->setName($this->fontName);//设置字体
        self::$obj->getDefaultStyle()->getFont()->setSize($this->fontSize);//设置字体大小
        self::$obj->getDefaultStyle()->getFont()->setBold($this->fontBold);//设置是否加粗
    }
    private function __clone()
    {
        // TODO: Implement __clone() method.
    }


    /**
     * 加载入口
     * @param array $config
     * @return Excel|null
     */
    public static function load($config = [])
    {
        if (!(self::$instance instanceof self)) {
            self::$instance = new self($config);
        }
        return self::$instance;
    }

    /**
     * @param array $header 标题 支持单行多行 标准格式  [
     * [
     * 'A' => ['title'=>'','value' => '','width'=>'','merge_column'=>'B','merge_line'=>''],
     * 'C' => ['title'=>'','value' => '','width'=>'','merge_column'=>'','merge_line'=>''],
     * ],
     * [
     * 'A' => ['title'=>'','value' => '','width'=>'','merge_column'=>'','merge_line'=>''],
     * 'B' => ['title'=>'','value' => '','width'=>'','merge_column'=>'','merge_line'=>''],
     * 'C' => ['title'=>'','value' => '','width'=>'','merge_column'=>'','merge_line'=>'1'],
     * ]
     * ]
     * title：中文标题 value：英文标题 width 列宽 merge_column 合并列 merge_line 向上合并行数
     * @param int $sheetIndex 内置表 默认表一
     * @param string $sheetName 内置表名称 默认 Sheet1
     * @return bool
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function setHeader($header = [],$sheetIndex = 0,$sheetName = 'Sheet1')
    {
        $header = self::setHeaderFormat($header);
        self::setHeaderData($header);
        if (!$header) return false;
        if (self::$headerLine == 0) return false;
        foreach ($header as $line => $head) {
            foreach ($head as $column => $data) {
                self::$obj->createSheet($sheetIndex);
                $sheet = self::$obj->getActiveSheet();
                $sheet->setTitle($sheetName);
                self::$obj->setActiveSheetIndex($sheetIndex)->setCellValue( $column.$line, $data['title']);
                if(!empty($data['merge_column']) && in_array($data['merge_column'],self::$columnList)) {
                    //合并列
                    self::$obj->getActiveSheet()->mergeCells($column.$line.':'.$data['merge_column'].$line);
                }
                if (!empty($data['merge_line']) && is_numeric($data['merge_line'])) {
                    //合并行
                    $lastLine = $line-$data['merge_line'];
                    if ($lastLine > 0) $lastLine = 1;//超过条数默认第一行开始
                    self::$obj->getActiveSheet()->mergeCells($column.$lastLine.':'.$column.$line);
                }
            }
        }
    }

    /**
     * 设置标题格式
     * @param $header
     * @return array|bool
     */
    private static function setHeaderFormat($header)
    {
        $data = [];
        $result = [];
        try {
            //处理数据
            foreach ($header as $key => $value) {
                if (!is_array($value)) {
                    //单行表头 格式为 title-value
                    $data[1][] = ['title'=>$key,'value' => $value,'width'=>self::$columnWidth,'merge_column'=>'','merge_line'=>''];
                    self::$headerLine = 1;
                } else {
                    self::$headerLine ++;
                    foreach ($value as $k =>$item) {
                        if (!is_array($item)) {
                            //多行表头 格式为 title-value
                            $data[self::$headerLine][] = ['title'=>$k,'value' => $value,'width'=>self::$columnWidth,'merge_column'=>'','merge_line'=>''];
                        } else {
                            //多行表头 标准格式
                            $data[self::$headerLine][] = [
                                'title'=>$item['title'],
                                'value' => $item['value'],
                                'width'=> isset($item['width']) ? $item['width'] : self::$columnWidth,
                                'merge_column'=>isset($item['merge_column']) ? $item['merge_column'] : '',
                                'merge_line'=>isset($item['merge_line']) ? $item['merge_line'] : ''
                            ];
                        }
                    }
                }
            }
            //将数字键名改成英文字母键名
            foreach ($data as $key => $value) {
                foreach ($value as $k => $column) {
                    if (!in_array($k,self::$columnList)) {
                        $result[self::$columnList[$k]] = $column;
                        continue;
                    }
                    $result[$k] = $column;
                }
            }
        }catch (\Throwable $exception) {
            return false;
        }
        return $result;
    }

    /**
     * 设置标题数据
     * @param $header
     */
    private static function setHeaderData($header)
    {
        foreach ($header as $line => $head) {
            foreach ($head as $column => $data) {
                self::$header[$column] = $data['value'];
                self::$obj->getActiveSheet()->getColumnDimension($column)->setWidth($data['width']);
            }
        }
    }

    /**
     * @param array $data  格式 [['value' => 'data']]
     * @param int $sheetIndex 内置表 默认表一
     */
    public static function setData($data,$sheetIndex = 0)
    {
        $startLine = self::$headerLine +1;
        foreach ($data as $lineData) {
            foreach ($lineData as $column => $columnData) {
                $column = array_search($column,self::$header);
                self::$obj->setActiveSheetIndex($sheetIndex)->setCellValue( $column.$startLine, $columnData);
            }
            $startLine++;
        }
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
    }

}