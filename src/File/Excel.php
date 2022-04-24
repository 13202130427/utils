<?php

namespace Uroad\Utils\File;



class Excel
{
//    private $example = [
//        [
//            'A' => ['title'=>'','value' => '','width'=>'','merge_column'=>'B'],
//            'C' => ['title'=>'','value' => '','width'=>''],
//        ],
//        [
//            'A' => ['title'=>'','value' => '','width'=>''],
//            'B' => ['title'=>'','value' => '','width'=>''],
//            'C' => ['title'=>'','value' => '','width'=>''],
//        ]
//    ];

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
        self::$obj = new \PHPExcel();
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

    public static function setHeader($header = [])
    {
        $header = self::setHeaderFormat($header);
        self::setHeaderData($header);
        if (!$header) return false;
        if (self::$headerLine == 0) return false;
        foreach ($header as $line => $head) {
            foreach ($head as $column => $data) {
                self::$obj->setActiveSheetIndex(0)->setCellValue( $column.$line, $data['title']);
                if(!empty($data['merge_column']) && in_array($data['merge_column'],self::$columnList)) {
                    //合并列
                    self::$obj->getActiveSheet()->mergeCells($column.$line.':'.$data['merge_column']);
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
                    $data[1][] = ['title'=>$key,'value' => $value,'width'=>self::$columnWidth,'merge_column'=>''];
                    self::$headerLine = 1;
                } else {
                    self::$headerLine ++;
                    foreach ($value as $k =>$item) {
                        if (!is_array($item)) {
                            //多行表头 格式为 title-value
                            $data[self::$headerLine][] = ['title'=>$k,'value' => $value,'width'=>self::$columnWidth,'merge_column'=>''];
                        } else {
                            //多行表头 标准格式
                            $data[self::$headerLine][] = [
                                'title'=>$item['title'],
                                'value' => $item['value'],
                                'width'=> isset($item['width']) ? $item['width'] : '',
                                'merge_column'=>isset($item['merge_column']) ? $item['merge_column'] : ''
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
     * @throws \PHPExcel_Exception
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
     * @param $data  [['value' => 'data']]
     */
    public static function setData($data)
    {
        $startLine = self::$headerLine +1;
        foreach ($data as $lineData) {
            foreach ($lineData as $column => $columnData) {
                $column = array_search($column,self::$header);
                self::$obj->setActiveSheetIndex(0)->setCellValue( $column.$startLine, $columnData);
            }
            $startLine++;
        }
    }

    public function generate($path ='',$format = 'Excel2007')
    {
        if (self::$setData && self::$setHeader) return false;
        if ($format == 'Excel2007') {
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        } elseif ($format == 'Excel5') {
            header('Content-Type: application/vnd.ms-excel');
        }
        self::$obj->setActiveSheetIndex(0);
        ob_end_clean();//清除缓冲区,避免乱码
        $objWriter = \PHPExcel_IOFactory::createWriter(self::$obj,$format);
        if (!empty($path)) {
            $objWriter->save($path);
            return $path;
        }
        $objWriter->save('php://output');
    }

}